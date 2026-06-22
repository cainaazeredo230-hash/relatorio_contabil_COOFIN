from playwright.sync_api import sync_playwright

URL = "https://siafe2-flexvision.fazenda.rj.gov.br/Flexvision/"


def fazer_login(playwright, cpf, senha):
    browser = playwright.chromium.launch(
        headless=False,
        channel="msedge"
    )

    context = browser.new_context(
        accept_downloads=True
    )

    page = context.new_page()

    page.goto(URL)
    page.wait_for_load_state("networkidle")

    page.locator("input[type='text']").first.fill(cpf)
    page.locator("input[type='password']").first.fill(senha)
    page.get_by_role("button").first.click()

    page.wait_for_timeout(8000)

    return browser, context, page