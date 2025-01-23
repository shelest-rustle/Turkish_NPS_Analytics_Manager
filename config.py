import logging


LOGGER = logging.getLogger(__name__)
HANDLER = logging.StreamHandler()
LOGGER.addHandler(HANDLER)

logging.basicConfig(level=logging.INFO)

UNITS = ["Kargo teslimat", "Ürün kurulumu", "Kurulum için zamanında Sonuç", "Ürün tanıtımı", "Servis tutum Sonuç", "Üründen mennunihyet Sonuç"]
