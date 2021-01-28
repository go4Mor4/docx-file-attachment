# Image base in Python 3.8.2
FROM python:3.8.2

# Create dir code
WORKDIR /code

# Copy folders to base path
COPY ./input /code/input
COPY ./logs /code/logs
COPY ./output /code/output
COPY ./src /code/src

# Copy file to base path
COPY __main__.py /code/__main__.py
COPY requirements.txt /code/requirements.txt

# Install Requirements
RUN pip install --upgrade pip
RUN pip install -r requirements.txt

# Create auxiliar folders
RUN mkdir export

# Expose Port 9001
EXPOSE 9001

# Run code
CMD ["python","."]