FROM node:20-alpine

WORKDIR /app

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm install --production

# Copy source code
COPY dynamicFetch.js ./

# Create a directory for persistent data
RUN mkdir -p /app/data

# Ensure the log file is pointing to the volume path
# We'll handle the path remapping in the code or via env
ENV AUDIT_LOG_PATH=/app/data/m365_audit_logs.json

EXPOSE 3001

CMD ["node", "dynamicFetch.js"]
