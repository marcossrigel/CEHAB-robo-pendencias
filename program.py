from seleniumbase import SB
import os
import re
import json
import time
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
COL_OBJETO = "OBJETO"

SEI_LOGIN_URL = "https://sei.pe.gov.br/sip/login.php?sigla_orgao_sistema=GOVPE&sigla_sistema=SEI"
XP_USUARIO = '//*[@id="txtUsuario"]'
XP_SENHA = '//*[@id="pwdSenha"]'
CSS_BTN_ACESSAR = '#sbmAcessar'
CSS_SELECT_ORGAO = "#selOrgao"
XP_TXT_PESQUISA_RAPIDA = '//*[@id="txtPesquisaRapida"]'
XP_BTN_LUPA = '//*[@id="spnInfraUnidade"]/img'

ROMAN_RE = re.compile(r"^(?=[IVXLCDM]+$)M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$")
REGEX_SEI = r"\d{7,}\.\d+\/\d{4}-\d+"
SEI_RE = re.compile(REGEX_SEI)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(BASE_DIR, "downloaded_files")
MAP_JSON = os.path.join(OUT_DIR, "sei_last_doc_map.json")

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


def fetch_seis_from_sheet_api() -> tuple[list[str], dict[str, str], dict[str, str]]:
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
    sei_to_objeto = {}

    for r in rows:
        status = normalize(str(r.get(COL_STATUS, "")))
        if "CONCLUÍDO" in status:
            continue

        raw = str(r.get(COL_SEI, "")).strip()
        sei = pick_last_sei_from_cell(raw)
        if not sei:
            continue

        dest = (str(r.get(COL_DEST, "")) or "").strip() or "—"
        objeto = (str(r.get(COL_OBJETO, "")) or "").strip() or "—"

        seis.append(sei)

        if sei not in sei_to_dest:
            sei_to_dest[sei] = dest

        if sei not in sei_to_objeto:
            sei_to_objeto[sei] = objeto

    uniq = list(dict.fromkeys(seis))
    return uniq, sei_to_dest, sei_to_objeto




def wait_until_not_visible_text(sb: SB, text: str, timeout: int = 15) -> None:
    end = time.time() + timeout
    while time.time() < end:
        try:
            if not sb.is_text_visible(text):
                return
        except Exception:
            return
        time.sleep(0.2)


def wait_for_tree_loaded(sb: SB, timeout: int = 15) -> None:
    end = time.time() + timeout
    while time.time() < end:
        try:
            spans = sb.find_elements("css selector", 'span[id^="span"]')
            icons = sb.find_elements("css selector", 'img[id^="icon"]')
            if spans or icons:
                return
        except Exception:
            pass
        time.sleep(0.2)


def is_roman(s: str) -> bool:
    s = (s or "").strip().upper()
    return bool(s) and bool(ROMAN_RE.match(s))


def sei_quick_search(sb: SB, sei: str) -> None:
    sb.wait_for_element_visible(XP_TXT_PESQUISA_RAPIDA, timeout=30)
    sb.clear(XP_TXT_PESQUISA_RAPIDA)
    sb.type(XP_TXT_PESQUISA_RAPIDA, sei)

    try:
        sb.click(XP_BTN_LUPA)
    except Exception:
        sb.js_click(XP_BTN_LUPA)

    wait_until_not_visible_text(sb, "Aguarde", timeout=15)
    sb.wait_for_ready_state_complete()


def find_tree_frame(sb: SB, timeout: int = 40):
    end = time.time() + timeout
    last_err = None

    while time.time() < end:
        try:
            sb.switch_to_default_content()
            frames = sb.find_elements("css selector", "iframe")
        except Exception as e:
            last_err = e
            time.sleep(0.2)
            continue

        for fr in frames:
            name = (fr.get_attribute("name") or "").strip()
            fid = (fr.get_attribute("id") or "").strip()
            key = name or fid or fr

            try:
                sb.switch_to_default_content()
                sb.switch_to_frame(key)

                spans = sb.find_elements("css selector", 'span[id^="span"]')
                for sp in spans[:80]:
                    txt = (sp.text or "").strip()
                    if txt and (is_roman(txt) or len(txt) > 3):
                        sb.switch_to_default_content()
                        return key
            except Exception as e:
                last_err = e
                continue

        time.sleep(0.2)

    sb.switch_to_default_content()
    raise RuntimeError(f"Não consegui localizar o iframe da árvore. Último erro: {last_err}")


def wait_for_roman_folders(sb: SB, timeout: int = 10) -> bool:
    end = time.time() + timeout
    while time.time() < end:
        spans = sb.find_elements("css selector", 'span[id^="span"]')
        for sp in spans:
            try:
                if not sp.is_displayed():
                    continue
                txt = (sp.text or "").strip()
                if is_roman(txt):
                    return True
            except Exception:
                pass
        time.sleep(0.2)
    return False


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
        return

    _, last_sp = romans[-1]
    sb.execute_script("arguments[0].scrollIntoView({block:'center'});", last_sp)

    parent = last_sp.find_element("xpath", "./..")
    imgs = parent.find_elements("css selector", "img")

    for img in imgs:
        try:
            src = (img.get_attribute("src") or "").lower()
            if "plus" in src or "expand" in src:
                img.click()
                wait_for_tree_loaded(sb, timeout=8)
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
        raise RuntimeError("Não achei arquivos visíveis na árvore.")

    return items


def wait_for_whatsapp_ready(sb: SB, timeout: int = 30) -> str:
    possiveis_caixas = [
        '//footer//*[@contenteditable="true"][@data-tab]',
        '//footer//*[@contenteditable="true"]',
        '//div[@contenteditable="true"][@role="textbox"]',
        '//div[@title="Digite uma mensagem"]',
        '//div[@title="Mensagem"]',
        '//p[@class and ancestor::*[@contenteditable="true"]]',
    ]

    end = time.time() + timeout
    while time.time() < end:
        for sel in possiveis_caixas:
            try:
                if sb.is_element_visible(sel):
                    return sel
            except Exception:
                pass
        time.sleep(0.5)

    raise RuntimeError("Não consegui localizar a caixa de mensagem do WhatsApp.")


def enviar_whatsapp(sb: SB, link_grupo: str, mensagem: str, timeout: int = 120):
    print("🔗 Abrindo link do grupo...")
    sb.open(link_grupo)
    sb.wait_for_ready_state_complete()

    time.sleep(2)

    possiveis_botoes_continuar = [
        '//*[@id="whatsapp-web-button"]',
        '//*[@id="whatsapp-web-button"]/span',
        '//a[@id="whatsapp-web-button"]',
        '//a[contains(@href, "web.whatsapp.com")]',
        '//span[contains(normalize-space(.), "Continuar para o WhatsApp Web")]',
        '//a[contains(., "Continuar para o WhatsApp Web")]',
    ]

    clicou_continuar = False

    for sel in possiveis_botoes_continuar:
        try:
            if sb.is_element_visible(sel, timeout=4):
                print(f"🖱️ Clicando em: {sel}")
                try:
                    sb.click(sel)
                except Exception:
                    sb.js_click(sel)
                clicou_continuar = True
                break
        except Exception:
            pass

    if not clicou_continuar:
        raise RuntimeError("Não consegui clicar em 'Continuar para o WhatsApp Web'.")

    time.sleep(4)

    try:
        sb.switch_to_window(-1)
    except Exception:
        pass

    end = time.time() + timeout
    while time.time() < end:
        try:
            url_atual = sb.get_current_url()
            if "web.whatsapp.com" in url_atual:
                break
        except Exception:
            pass
        time.sleep(0.5)

    sb.wait_for_ready_state_complete()

    caixa_usada = wait_for_whatsapp_ready(sb, timeout=timeout)

    print(f"✅ Caixa encontrada: {caixa_usada}")

    try:
        sb.click(caixa_usada)
    except Exception:
        sb.js_click(caixa_usada)

    pyperclip.copy(mensagem)

    enviado_texto = False
    try:
        el = sb.find_element(caixa_usada)
        el.send_keys(Keys.CONTROL, "v")
        enviado_texto = True
    except Exception:
        pass

    if not enviado_texto:
        try:
            sb.type(caixa_usada, mensagem)
            enviado_texto = True
        except Exception:
            pass

    if not enviado_texto:
        raise RuntimeError("Não consegui preencher a caixa de mensagem no WhatsApp.")

    time.sleep(1)

    botoes_enviar = [
        '//button[@aria-label="Enviar"]',
        '//span[@data-icon="send"]/ancestor::button',
        '//button[.//span[@data-icon="send"]]',
        '//div[@role="button"]//span[@data-icon="send"]/ancestor::div[@role="button"]',
    ]

    enviado = False
    for botao in botoes_enviar:
        try:
            if sb.is_element_visible(botao, timeout=3):
                sb.click(botao)
                enviado = True
                break
        except Exception:
            pass

    if not enviado:
        try:
            el = sb.find_element(caixa_usada)
            el.send_keys(Keys.ENTER)
            enviado = True
        except Exception:
            pass

    if enviado:
        print("📨 Mensagem enviada no grupo!")
    else:
        raise RuntimeError("Não consegui enviar a mensagem no WhatsApp.")


def main():
    os.makedirs(OUT_DIR, exist_ok=True)

    seis, sei_to_dest, sei_to_objeto = fetch_seis_from_sheet_api()
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

        sb.open(SEI_LOGIN_URL)
        sb.wait_for_ready_state_complete()

        if not sb.is_element_visible(XP_TXT_PESQUISA_RAPIDA):
            sb.wait_for_element_visible(XP_USUARIO, timeout=30)
            sb.type(XP_USUARIO, sei_user)

            sb.wait_for_element_visible(XP_SENHA, timeout=30)
            sb.type(XP_SENHA, sei_pass)

            sb.wait_for_element_visible(CSS_SELECT_ORGAO, timeout=30)
            sb.select_option_by_text(CSS_SELECT_ORGAO, "CEHAB")

            sb.wait_for_element_visible(CSS_BTN_ACESSAR, timeout=30)
            sb.click(CSS_BTN_ACESSAR)

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
        tree_frame = find_tree_frame(sb, timeout=40)

        for idx, sei in enumerate(seis, start=1):
            print(f"\n[{idx}/{len(seis)}] 🔎 SEI: {sei}")
            print("   👤 Destinatário:", sei_to_dest.get(sei, "—"))
            print("   📌 Objeto      :", sei_to_objeto.get(sei, "—"))

            try:
                sei_quick_search(sb, sei)

                sb.switch_to_default_content()
                sb.switch_to_frame(tree_frame)

                achou_romano = wait_for_roman_folders(sb, timeout=8)
                if achou_romano:
                    expand_last_roman_folder(sb)

                wait_for_tree_loaded(sb, timeout=8)
                items = get_visible_files_in_tree(sb)
                textos = [t for _, t in items]

                anterior = (new_map.get(sei) or "").strip()

                ultimo_txt = ""
                novos = []

                if items:
                    ultimo_txt = items[-1][1]

                    if not anterior:
                        novos = [items[-1]]
                    elif anterior == ultimo_txt:
                        novos = []
                    elif anterior in textos:
                        idx_prev = textos.index(anterior)
                        novos = items[idx_prev + 1:]
                    else:
                        novos = [items[-1]]

                novos_txts = [txt for _, txt in novos]
                qtd_novos = len(novos)
                mudou = qtd_novos > 0

                if mudou and novos:
                    mudancas[sei] = {
                        "qtd_novos": qtd_novos,
                        "ultimo": ultimo_txt,
                        "novos": novos_txts,
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
                save_map(new_map)

            except Exception as e:
                print("   ❌ Erro neste SEI:", repr(e))
            finally:
                sb.switch_to_default_content()

        linhas = []
        data_msg = datetime.now().strftime("%d/%m/%Y")

        linhas.append(f"⚠️ Solicitações (Pendências) SEPLAG/SEFAZ --> Acompanhamento 2026 ⚠️ dia {data_msg}")
        linhas.append("📌 SEIs com novos documentos:")
        linhas.append("------------------------------")

        if not mudancas:
            linhas.append("Nenhum SEI mudou ✅")
        else:
            for sei_k, info in mudancas.items():
                dest = sei_to_dest.get(sei_k, "—")
                objeto = sei_to_objeto.get(sei_k, "—")

                linhas.append(f"{sei_k} - {dest}")
                linhas.append(f"Objeto: {objeto}")
                for doc in info.get("novos", []):
                    linhas.append(f"-> {doc}")
                linhas.append("")

        mensagem_final = "\n".join(linhas)

        print("\n==============================")
        print(mensagem_final)
        print("==============================")

        # salva o mapa ANTES de qualquer envio
        save_map(new_map)

        if mudancas:
            try:
                enviar_whatsapp(sb, "https://chat.whatsapp.com/Dve4KOqA55x0Mu56AqD4Ad", mensagem_final)
            except Exception as e:
                print(f"⚠️ Erro ao enviar no WhatsApp: {e}")

    print("\n✅ Finalizado com sucesso!")
    input("\n👉 Pressione ENTER para fechar o terminal...")


if __name__ == "__main__":
    main()
