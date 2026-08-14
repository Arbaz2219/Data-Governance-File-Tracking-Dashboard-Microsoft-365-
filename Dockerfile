FROM node:20-alpine

WORKDIR /app

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm install --production

# Copy source code
COPY dynamicFetch.js ./
# Seed list for a fresh data volume; the live list lives on the volume itself
COPY authorized_users.json ./

# Create a directory for persistent data
RUN mkdir -p /app/data

# Ensure the log file is pointing to the volume path
# We'll handle the path remapping in the code or via env
ENV AUDIT_LOG_PATH=/app/data/m365_audit_logs.json
# Keep mutable state on the volume too, otherwise it is lost on every redeploy
ENV AUTHORIZED_USERS_PATH=/app/data/authorized_users.json
ENV WEB_LOG_PATH=/app/data/web_activity_logs.json

EXPOSE 3001

CMD ["node", "dynamicFetch.js"]
