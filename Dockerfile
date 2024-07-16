# Usa una imagen base oficial de Python
FROM python:3.9

# Establece el directorio de trabajo en /app
WORKDIR /app

# Copia el archivo requirements.txt en el directorio de trabajo
COPY requirements.txt ./

# Instala las dependencias
RUN pip install --no-cache-dir -r requirements.txt

RUN apt-get update
RUN apt-get install wkhtmltopdf texlive-xetex -y

# Copia todo el contenido de la carpeta actual en el directorio de trabajo
COPY . .

# Expone el puerto que usa Streamlit
EXPOSE 8501

# Define el comando por defecto para ejecutar el contenedor
CMD ["streamlit", "run", "pdfform.py", "--server.port=8501", "--server.address=0.0.0.0"]
