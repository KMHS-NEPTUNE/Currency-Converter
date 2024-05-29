# Use alpine linux as base image
FROM alpine:latest

# Set environment variables for configuration
ENV APP_HOME /app
ENV PORT 8000

# Set working directory
WORKDIR $APP_HOME

# Install Python3 and pip
RUN apk add --no-cache python3 py3-pip

# Copy requirements.txt to the working directory
COPY requirements.txt .

# Install Python dependencies
RUN pip3 install -r requirements.txt --break-system-packages

# Expose the port
EXPOSE $PORT

# Copy the FastAPI app to the working directory
COPY . .

# Run FastAPI
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]

# End of Dockerfile
