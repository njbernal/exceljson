FROM python:3.10-buster
WORKDIR /ExcelJSON
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
CMD ["python", "server.py"]