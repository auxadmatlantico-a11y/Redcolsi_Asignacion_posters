import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto(
        "https://www.redcolsiplataforma.org/login.php?idPoison=ThePrincess6a020eb4e91039.280169151778519732")
    page.get_by_role("textbox", name="Documento").click()
    page.get_by_role("textbox", name="Documento").fill("1143453344")
    page.get_by_role("textbox", name="Contraseña").click()
    page.get_by_role("textbox", name="Contraseña").fill("SemillaPCA2024")
    page.get_by_role("button", name="Inicia Sesión").click()
    page.get_by_role("link", name=":: Ingresar ::.").nth(1).click()
    page.get_by_role("link", name="   Eventos").click()
    page.get_by_role("link", name="2").click()
    page.locator(
        "tr:nth-child(11) > td:nth-child(8) > .dropdown > #dropdownMenuButton").click()
    page.get_by_role("link", name="   Editar").click()
    page.get_by_role("tab", name="Programación").click()
    # Aqui en programacion de fecha se escoge con la rueda del raton
    page.locator("#ProgramacionFecha").click()
    page.get_by_role("button", name="Ok").click()
    page.locator("#ProgramacionHoraIni").click()
    page.locator("#ProgramacionHoraIni").fill("10:00")
    page.locator("#ProgramacionHoraFin").click()
    page.locator("#ProgramacionHoraFin").fill("13:00")
    page.get_by_role("textbox", name="Lugar programado").click()
    page.get_by_role("textbox", name="Lugar programado").fill(
        "Unilibre sede norte")
    page.get_by_role("textbox", name="Descripción de la programación").click()
    page.get_by_role(
        "textbox", name="Descripción de la programación").fill("prueba")
    page.get_by_role("textbox", name="Rango inicial poster").click()

    # Se escoge segun el excel
    page.get_by_role("textbox", name="Rango inicial poster").fill("1")

    # se escoge segun el titulo del proyecto en el excel y toca marcar el checkBox que corresponde al titulo del proyecto
    page.locator("tr:nth-child(602) > td > .ProgramaChk").check()
    page.get_by_role("button", name="Agregar programación").click()
    page.get_by_role("button", name="Aceptar").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
