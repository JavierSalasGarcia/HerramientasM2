import qrcode
from PIL import Image
import os

# Crear la carpeta si no existe
carpeta_destino = "C:\\codigoQR"
if not os.path.exists(carpeta_destino):
    os.makedirs(carpeta_destino)

# Pedir la URL al usuario
url = input("Introduce la dirección URL: ")

# Generar el código QR
qr = qrcode.QRCode(
    version=1,
    error_correction=qrcode.constants.ERROR_CORRECT_L,
    box_size=10,
    border=4,
)
qr.add_data(url)
qr.make(fit=True)

# Crear una imagen a partir del código QR
img = qr.make_image(fill_color="black", back_color="white")

# Guardar la imagen en el archivo
ruta_archivo = os.path.join(carpeta_destino, "codigoQR.jpg")
img.save(ruta_archivo)

print(f"El código QR se ha guardado en: {ruta_archivo}")
