FROM python:3.9.4-slim

COPY requirements.txt /app/
WORKDIR /app

RUN pip install --upgrade pip \
    &&  pip install --trusted-host pypi.python.org --requirement requirements.txt

COPY ClointFusion.py /app

CMD ["python", "ClointFusion.py"]