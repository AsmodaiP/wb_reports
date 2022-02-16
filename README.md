# Описание проекта

Проект создан для обработки файло в отчета Wildberries и занесения информации:

1. О количетсве продаж
2. О сумме "Вайлдберриз реализовал Товар (Пр)"
3. О сумме "Вознаграждение Вайлдберриз без НДС"

## Инструкция по работе

1. Пришлите телеграм-боту документ отчета Wildberries.
2. Пришлите боту название листа, на которые нужно вносить данные.
3. В случае если каких-то артикулов не будет в таблице бот пришлет сообщение об их отстутсвии. В таком случае нужно добавить информацию об этих артикулах в таблицу.
4. Внесите сумму возвратов, присланную ботом, в соответствующую ячейку таблицы.

## Деплой

1. Скопируйте репозиторий

   ```bash
    git clone https://github.com/AsmodaiP/wb_reports
   ```

2. Перейдите в клонированный репозиторий  и создайте виртуальное окружение.

    ```bash
    python3 -m venv venv
    ```

3. Активируйте виртуальное окружени и установите зависимости

    ```bash
    source venv/bin/activate
    ```

    ```bash
    pip install -r requirements.txt
    ```

4. Создайте файл .env и наполните его по следующему образцу:

    ```bash
    TELEGRAM_TOKEN=%токен телеграм-бота%
    SPREADSHEET_ID=%id таблицы, в которую нужно вносить изменения%
    ```

5. Поместите файл credentials_service.json в папку с проектом. Получить этот файл можно действуя согласно [инструкциям самого гугла](https://developers.google.com/workspace/guides/create-credentials)

6. Запустите bot.py

    ```bash
    python bot.py
    ```

В случае, если вы хотите, чтобы бот работал фоном, то следует запускать бот через команду ```nohup python bot.py &```

## Постоянная работа
Для постоянной работы нужно выполнить

```bash
sudo nano
lib/systemd/system/bot.service
```

И заполнить файл  следующим образом

```bash
[Unit]
Description=bot for rocket
After=network.target

[Service]
User=root
EnviromentFile=/etc/environment
ExecStart=путь_до_виртуального_окружения/bin/python bot.py 
ExecReload=путь_до_виртуального_окружения/bin/python bot.py 
WorkingDirectory=путь_до_проекта
KillMode=process
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target

```

После чего нужно выполнить

```bash
sudo  systemctl enable bot
sudo systemctl daemon-reload
```
