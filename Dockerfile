FROM python:3.6

WORKDIR /bot

COPY . /bot
COPY Pipfile Pipfile.lock ./


ENV TZ=Asia/Ho_Chi_Minh
RUN ln -snf /usr/share/zoneinfo/$TZ /etc/localtime && echo $TZ > /etc/timezone
RUN pip install -U pipenv
RUN pipenv install --system
RUN chmod a+rwx .tokens
# CMD ["python", "fetch_bot.py"]