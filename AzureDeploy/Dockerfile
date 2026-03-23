# ---- build stage ----
FROM python:3.12-slim AS build
WORKDIR /build
COPY requirements.txt .
RUN pip install --no-cache-dir --target=/build/deps -r requirements.txt

# ---- runtime stage ----
FROM python:3.12-slim
WORKDIR /app

# Installed dependencies from build stage
COPY --from=build /build/deps /usr/local/lib/python3.12/site-packages

# Application code
COPY TeamsCommunication/ ./TeamsCommunication/
COPY md_to_docx/ ./md_to_docx/
COPY DocTemplates/ ./DocTemplates/
COPY requirements.txt .
COPY env .

EXPOSE 8000

CMD ["python", "-m", "TeamsCommunication.app"]
