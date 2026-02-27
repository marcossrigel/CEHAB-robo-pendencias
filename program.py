from seleniumbase import SB
import os
import re
import time
import json
import csv
from datetime import datetime

import pyperclip
from selenium.webdriver.common.keys import Keys
import gspread
from oauth2client.service_account import ServiceAccountCredentials

CRED_JSON = r"c:\temp\ws-python-pendencias\credenciais_sheets.json"
SHEET_ID = "1CFI3282Mx7MDw13RK5Vq7cA0BJxKsN3h7pwjBzFVogA"
WORKSHEET_TITLE = "Acompanhamento 2026"

COL_SEI = "SEI"
COL_STATUS = "STATUS"
COL_DEST = "DESTINATÁRIO"

SEI_LOGIN_URL = "https://sei.pe.gov.br/sip/login.php?sigla_orgao_sistema=GOVPE&sigla_sistema=SEI"

XP_USUARIO = '//*[@id="txtUsuario"]'
XP_SENHA = '//*[@id="pwdSenha"]'
CSS_BTN_ACESSAR = '#sbmAcessar'
XP_BTN_ACESSAR = '//*[@id="sbmAcessar"]'
CSS_SELECT_ORGAO = "#selOrgao"
XP_TXT_PESQUISA_RAPIDA = '//*[@id="txtPesquisaRapida"]'
XP_BTN_LUPA = '//*[@id="spnInfraUnidade"]/img'

ROMAN_RE = re.compile(r"^(?=[IVXLCDM]+$)M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$")
REGEX_SEI = r"\d{7,}\.\d+\/\d{4}-\d+"
SEI_RE = re.compile(REGEX_SEI)


OUT_DIR = "downloaded_files"
MAP_JSON = os.path.join(OUT_DIR, "sei_last_doc_map.json")


def now_ts():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def normalize(s: str) -> str:
    return (s or "").strip().upper()


def safe_name(s: str) -> str:
    s = (s or "").strip()
    return re.sub(r"[^a-zA-Z0-9_.-]+", "_", s)[:120]


def load_map() -> dict:
    if not os.path.exists(MAP_JSON):
        return {}
    try:
        with open(MAP_JSON, "r", encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        return {}


def save_map(data: dict) -> None:
    os.makedirs(OUT_DIR, exist_ok=True)
    with open(MAP_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def pick_last_sei_from_cell(cell: str) -> str:

    text = (cell or "").strip()
    if not text:
        return ""
    matches = SEI_RE.findall(text)
    return matches[-1].strip() if matches else text


def fetch_seis_from_sheet_api() -> tuple[list[str], dict[str, str]]:

    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_name(CRED_JSON, scope)
    client = gspread.authorize(creds)

    sh = client.open_by_key(SHEET_ID)
    ws = sh.worksheet(WORKSHEET_TITLE)

    rows = ws.get_all_records()

    seis = []
    sei_to_dest = {}

    for r in rows:
        status = normalize(str(r.get(COL_STATUS, "")))
        if "CONCLUÍDO" in status:
            continue

        raw = str(r.get(COL_SEI, "")).strip()
        sei = pick_last_sei_from_cell(raw)
        if not sei:
            continue

        dest = (str(r.get(COL_DEST, "")) or "").strip()
        if not dest:
            dest = "—"

        seis.append(sei)

        if sei not in sei_to_dest:
            sei_to_dest[sei] = dest

    uniq = list(dict.fromkeys(seis))
    return uniq, sei_to_dest

def is_roman(s: str) -> bool:
    s = (s or "").strip().upper()
    return bool(s) and bool(ROMAN_RE.match(s))


def sei_quick_search(sb: SB, sei: str) -> None:
    sb.wait_for_element_visible(XP_TXT_PESQUISA_RAPIDA, timeout=40)
    sb.click(XP_TXT_PESQUISA_RAPIDA)
    sb.clear(XP_TXT_PESQUISA_RAPIDA)
    sb.type(XP_TXT_PESQUISA_RAPIDA, sei)
    sb.click(XP_BTN_LUPA)

    sb.sleep(3.0)

def wait_for_roman_folders(sb: SB, timeout: int = 12) -> bool:
    end = time.time() + timeout
    while time.time() < end:
        spans = sb.find_elements("css selector", 'span[id^="span"]')

        if spans:
            for sp in spans:
                try:
                    if not sp.is_displayed():
                        continue
                    txt = (sp.text or "").strip()
                    if is_roman(txt):
                        return True
                except Exception:
                    pass
        time.sleep(0.4)
    return False


def find_tree_frame(sb: SB, timeout: int = 60) -> str:
    end = time.time() + timeout
    last_err = None

    while time.time() < end:
        try:
            sb.switch_to_default_content()
            frames = sb.find_elements("css selector", "iframe")
        except Exception as e:
            last_err = e
            time.sleep(0.5)
            continue

        for fr in frames:
            name = (fr.get_attribute("name") or "").strip()
            fid = (fr.get_attribute("id") or "").strip()
            key = name or fid
            if not key:
                continue
            try:
                sb.switch_to_default_content()
                sb.switch_to_frame(key)

                spans = sb.find_elements("css selector", 'span[id^="span"]')
                for sp in spans[:120]:
                    txt = (sp.text or "").strip()
                    if is_roman(txt):
                        sb.switch_to_default_content()
                        return key
            except Exception as e:
                last_err = e
                continue

        time.sleep(0.6)

    sb.switch_to_default_content()
    raise RuntimeError(f"Não consegui localizar o iframe da árvore automaticamente. Último erro: {last_err}")


def expand_last_roman_folder(sb: SB) -> None:
    spans = sb.find_elements("css selector", 'span[id^="span"]')
    romans = []
    for sp in spans:
        try:
            if not sp.is_displayed():
                continue
            txt = (sp.text or "").strip()
            if is_roman(txt):
                romans.append((txt, sp))
        except Exception:
            continue

    if not romans:
        raise RuntimeError("Não achei nenhuma pasta romana na árvore.")

    _, last_sp = romans[-1]
    sb.execute_script("arguments[0].scrollIntoView({block:'center'});", last_sp)
    sb.sleep(0.15)

    parent = last_sp.find_element("xpath", "./..")
    imgs = parent.find_elements("css selector", "img")
    for img in imgs:
        try:
            src = (img.get_attribute("src") or "").lower()
            if "plus" in src or "expand" in src:
                img.click()
                sb.sleep(0.6)
                return
        except Exception:
            pass


def get_visible_files_in_tree(sb: SB) -> list[tuple[str, str]]:
    icons = sb.find_elements("css selector", 'img[id^="icon"]')
    items = []

    for ic in icons:
        try:
            if not ic.is_displayed():
                continue

            icon_id = (ic.get_attribute("id") or "").strip()
            if not icon_id.startswith("icon"):
                continue

            num = icon_id.replace("icon", "").strip()
            if not num.isdigit():
                continue

            span_id = f"span{num}"
            sp = sb.find_element("css selector", f"span#{span_id}")
            if not sp.is_displayed():
                continue

            txt = (sp.text or "").strip()
            if not txt:
                continue

            items.append((num, txt))
        except Exception:
            continue

    if not items:
        raise RuntimeError("Não achei arquivos visíveis (img#icon... + span#span...).")

    return items


def click_papel_azul_do_item(sb: SB, num: str) -> None:
    xp_icon = f'//*[@id="icon{num}"]'
    sb.wait_for_element_visible(xp_icon, timeout=15)
    sb.scroll_to(xp_icon)
    try:
        sb.js_click(xp_icon)
    except Exception:
        sb.click(xp_icon)
    sb.sleep(0.12)


def open_doc(sb: SB, num: str) -> None:
    xp_span = f'//*[@id="span{num}"]'
    sb.wait_for_element_visible(xp_span, timeout=15)
    sb.scroll_to(xp_span)
    try:
        sb.js_click(xp_span)
    except Exception:
        sb.click(xp_span)


def save_screenshot(sb: SB, folder: str, prefix: str) -> str:
    os.makedirs(folder, exist_ok=True)
    filename = f"{prefix}_{now_ts()}.png"
    path = os.path.join(folder, filename)
    sb.save_screenshot(path)
    return path


def enviar_whatsapp(sb: SB, link_grupo: str, mensagem: str, timeout: int = 120):
    print("🔗 Abrindo link do grupo...")
    sb.open(link_grupo)

    btn_continuar = '//*[@id="whatsapp-web-button"]/span'
    sb.wait_for_element_visible(btn_continuar, timeout=timeout)
    sb.click(btn_continuar)
    sb.sleep(2)

    try:
        sb.switch_to_window(-1)
    except Exception:
        pass

    caixa_msg = '//*[@id="main"]/footer/div[1]/div/span/div/div/div/div[3]/div[1]/p'
    sb.wait_for_element_visible(caixa_msg, timeout=timeout)
    sb.click(caixa_msg)
    sb.sleep(0.3)

    pyperclip.copy(mensagem)
    el = sb.find_element(caixa_msg)
    el.send_keys(Keys.CONTROL, "v")

    sb.sleep(1.5)
    el.send_keys(Keys.ENTER)
    sb.sleep(3)
    print("📨 Mensagem enviada no grupo!")


def main():
    os.makedirs(OUT_DIR, exist_ok=True)

    pasta_hoje = os.path.join(OUT_DIR, datetime.now().strftime("%Y-%m-%d"))
    os.makedirs(pasta_hoje, exist_ok=True)

    seis, sei_to_dest = fetch_seis_from_sheet_api()
    if not seis:
        print("⚠️ Nenhum SEI encontrado (ou todos estão CONCLUÍDO).")
        return

    print(f"📄 SEIs válidos (não concluídos): {len(seis)}")

    old_map = load_map()
    new_map = dict(old_map)

    mudancas = {} 

    sei_user = os.getenv("SEI_USER", "marcos.rigel")
    sei_pass = os.getenv("SEI_PASS", "Abc123!@")

    with SB(
        uc=False,
        headless=False,
        user_data_dir="C:/temp/chrome_profile_whatsapp"
    ) as sb:
        sb.maximize_window()

        sb.open(SEI_LOGIN_URL)
        sb.wait_for_ready_state_complete()

        if sb.is_element_visible(XP_TXT_PESQUISA_RAPIDA):
            print("✅ Já estava logado (pulando login).")
        else:
            sb.wait_for_element_visible(XP_USUARIO, timeout=30)
            sb.type(XP_USUARIO, sei_user)

            sb.wait_for_element_visible(XP_SENHA, timeout=30)
            sb.type(XP_SENHA, sei_pass)

            sb.wait_for_element_visible(CSS_SELECT_ORGAO, timeout=30)
            sb.select_option_by_text(CSS_SELECT_ORGAO, "CEHAB")
            sb.sleep(0.5)

            sb.wait_for_element_visible(CSS_BTN_ACESSAR, timeout=30)
            sb.click(CSS_BTN_ACESSAR)
            sb.sleep(1.5)

        try:
            sb.accept_alert(timeout=2)
        except Exception:
            pass

        try:
            sb.switch_to_window(-1)
        except Exception:
            pass

        sb.wait_for_element_visible(XP_TXT_PESQUISA_RAPIDA, timeout=60)

        sei_quick_search(sb, seis[0])
        tree_frame = find_tree_frame(sb, timeout=80)

        for idx, sei in enumerate(seis, start=1):
            print(f"\n[{idx}/{len(seis)}] 🔎 SEI: {sei}")
            print("   👤 Destinatário:", sei_to_dest.get(sei, "—"))

            try:
                sei_quick_search(sb, sei)

                sb.switch_to_default_content()
                sb.switch_to_frame(tree_frame)
                achou_romano = wait_for_roman_folders(sb, timeout=14)

                if achou_romano:
                    expand_last_roman_folder(sb)
                    sb.sleep(0.8)
                else:
                    sb.sleep(1.2)

                items = get_visible_files_in_tree(sb)
                textos = [t for _, t in items]

                anterior = (old_map.get(sei) or "").strip()

                if anterior and anterior in textos:
                    idx_prev = textos.index(anterior)
                    novos = items[idx_prev + 1:]
                else:
                    novos = [items[-1]]

                novos_txts = [txt for _, txt in novos]
                ultimo_txt = items[-1][1]

                qtd_novos = len(novos)
                mudou = (qtd_novos > 0 and (not anterior or ultimo_txt != anterior))

                screenshot_paths = []
                if mudou and novos:
                    for num_item, txt_item in novos:
                        click_papel_azul_do_item(sb, num_item)
                        open_doc(sb, num_item)
                        sb.sleep(1.2)

                        prefix = f"sei_{safe_name(sei)}_{safe_name(txt_item)[:50]}"
                        shot = save_screenshot(sb, folder=pasta_hoje, prefix=prefix)
                        screenshot_paths.append(shot)

                    mudancas[sei] = {
                        "qtd_novos": qtd_novos,
                        "ultimo": ultimo_txt,
                        "novos": novos_txts,
                        "screens": screenshot_paths,
                    }

                print("   ✅ Último doc:", ultimo_txt)
                if anterior:
                    print("   🗂️  Anterior :", anterior)
                print("   🆕 Novos docs:", qtd_novos)
                if novos_txts:
                    for t in novos_txts:
                        print("   ->", t)
                print("   🔁 Mudou?    :", "SIM" if mudou else "NÃO")

                new_map[sei] = ultimo_txt

            except Exception as e:
                print("   ❌ Erro neste SEI:", repr(e))
            finally:
                sb.switch_to_default_content()

        linhas = []
        linhas.append("⚠️ Solicitações (Pendências) SEPLAG/SEFAZ --> Acompanhamento 2026 ⚠️")
        linhas.append("📌 SEIs com novos documentos:")
        linhas.append("------------------------------")

        if not mudancas:
            linhas.append("Nenhum SEI mudou ✅")
        else:
            for sei_k, info in mudancas.items():
                dest = sei_to_dest.get(sei_k, "—")
                linhas.append(f"{sei_k} - {dest}")
                for doc in info.get("novos", []):
                    linhas.append(f"-> {doc}")
                linhas.append("")

        mensagem_final = "\n".join(linhas)

        print("\n==============================")
        print(mensagem_final)
        print("==============================")

        if mudancas:
            enviar_whatsapp(sb, "https://chat.whatsapp.com/Dve4KOqA55x0Mu56AqD4Ad", mensagem_final)
            sb.sleep(10)

    save_map(new_map)

    print("\n✅ Finalizado com sucesso!")
    input("\n👉 Pressione ENTER para fechar o terminal...")


if __name__ == "__main__":
    main()
