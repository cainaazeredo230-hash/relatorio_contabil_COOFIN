import re
from playwright.sync_api import Page
import logging

logging.basicConfig(
    level=logging.INFO,
    filename="app.log",
    filemode="a",
    format="%(asctime)s - %(levelname)s - %(message)s"
)

def baixar_relatorio(page: Page,
                     usuario="Cainã",
                     codigo="082831",
                     arquivo="relatorio_automatico.xlsx"):

    page.get_by_role("button", name=" Consultas").click()
    page.goto(
        "https://siafe2-flexvision.fazenda.rj.gov.br/Flexvision/#!consultas_"
    )

    page.get_by_role("treegrid").get_by_text(usuario).click()
    page.get_by_role("gridcell", name=codigo).click()
    page.get_by_role("button").filter(
        has_text=re.compile(r"^$")
    ).click()

    page.get_by_text("").wait_for(timeout=1800000)
    page.get_by_text("").click()

    with page.expect_download(timeout=60000) as download_info:
        page.get_by_role("button", name=" Excel").click()

    download = download_info.value
    download.save_as(arquivo)

    logging.info(f"Arquivo salvo em: {arquivo}")