FROM python:3.11-slim

# Install system deps if needed later (e.g. for Pillow/fonts)
RUN apt-get update && apt-get install -y --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install dependencies first (better layering)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the code
COPY . .

# Render will set PORT, but default to 8000 if not
ENV PORT=8000

EXPOSE 8000

# Start MCP server over HTTP, binding to 0.0.0.0 and $PORT
CMD ["sh", "-c", "python ppt_mcp_server.py --transport http --port ${PORT:-8000}"]

