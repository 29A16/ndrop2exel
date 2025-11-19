# select starting image
FROM python:3.11-slim

# Create user name and home directory variables. 
# The variables are later used as $USER and $HOME. 
ENV USER=username
ENV HOME=/home/$USER

# Add user to system
RUN useradd -m -u 1000 $USER

# Set working directory (this is where the code should go)
WORKDIR $HOME/app

# Update system and install dependencies including libgxps-utils for XPS conversion
RUN apt-get update && apt-get install --no-install-recommends -y \
    build-essential \
    libgxps-utils \
    openjdk-21-jre \
    curl \
    && rm -rf /var/lib/apt/lists/*


# Copy application files
COPY app/app.py $HOME/app/app.py
COPY app/requirements.txt $HOME/app/requirements.txt


# Install Python packages listed in requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Change ownership of app directory to user
RUN chown -R $USER:$USER $HOME/app

USER $USER
EXPOSE 8501

HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health

ENTRYPOINT ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0", "--browser.gatherUsageStats=false"]