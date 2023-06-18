# Use an appropriate base image with Python installed
FROM python:3.9

# Set the working directory inside the container
WORKDIR /app

# Copy the Python application file to the container
COPY Auto_Trade_Active_CSV_Updated_0316(latest).py .

# Install the Python dependencies
RUN pip install --no-cache-dir grequests pandas openpyxl

# Specify the command to run your application
CMD ["python", "Auto_Trade_Active_CSV_Updated_0316(latest).py"]
