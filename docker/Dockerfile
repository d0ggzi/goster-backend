FROM python:3.11.5

ENV PYTHONDONTWRITEBYTECODE 1
ENV PYTHONUNBUFFERED 1

WORKDIR /opt/vkr-part

RUN pip install poetry==2.2.1

COPY src ./src
COPY migrations ./migrations
COPY pyproject.toml poetry.lock alembic.ini README.md ./

RUN poetry config virtualenvs.create false \
    && poetry install

CMD ["sh", "-c", "alembic upgrade head && uvicorn src.main:app --host 0.0.0.0 --port 8080"]

