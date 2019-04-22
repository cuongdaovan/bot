FROM python:3.6

WORKDIR /bot

COPY . /bot
COPY Pipfile Pipfile.lock ./


ENV TZ=Asia/Ho_Chi_Minh
RUN ln -snf /usr/share/zoneinfo/$TZ /etc/localtime && echo $TZ > /etc/timezone
RUN pip install -U pipenv
RUN pipenv install --system
RUN apt-get update && apt-get install -y vim

# CMD ["python", "fetch_bot.py"]