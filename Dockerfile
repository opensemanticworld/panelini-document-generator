# Install uv
FROM python:3.12-slim-trixie
COPY --from=ghcr.io/astral-sh/uv:latest /uv /bin/uv

ENV DEBIAN_FRONTEND noninteractive

RUN apt-get update && apt-get -y upgrade && \
    apt-get --no-install-recommends install libreoffice -y && \
    apt-get --no-install-recommends install libreoffice-java-common -y

# Change the working directory to the `app` directory
WORKDIR /app

# Copy the lockfile and `pyproject.toml` into the image
COPY uv.lock /app/uv.lock
COPY pyproject.toml /app/pyproject.toml

# Install dependencies
RUN uv sync --frozen --no-install-project

# Copy the project into the image
COPY . /app

# Sync the project
RUN uv sync --frozen

CMD [ "uv", "run", "panel", "serve", "src/document-generator/main.py", "--address", "0.0.0.0", "--port", "80", "--static-dirs", "assets=.venv/lib/python3.12/site-packages/panelini/assets", "--ico-path", ".venv/lib/python3.12/site-packages/panelini/assets/favicon.ico"]