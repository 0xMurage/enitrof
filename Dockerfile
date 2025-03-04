FROM python:3.12.9-alpine3.21

WORKDIR /app

COPY requirements.txt .

RUN pip install -r requirements.txt

COPY . .

CMD ["python","index:run"]