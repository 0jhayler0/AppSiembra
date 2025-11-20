# Use a base image with Node.js and Python
FROM node:18-bullseye

# Install Python and pip
RUN apt-get update && apt-get install -y python3 python3-pip && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy package.json and install Node.js dependencies
COPY backend/package*.json ./backend/
RUN cd backend && npm install

# Install Python dependencies
RUN pip3 install camelot-py[base] PyMuPDF pdfplumber openpyxl

# Copy the rest of the application
COPY . .

# Build the frontend
RUN cd backend && npm run build

# Expose the port
EXPOSE 5000

# Start the application
CMD ["node", "backend/server.js"]
