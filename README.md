# Бот для MS Teams для заполнения листов учета рабочего времени

Бот формирует лист учета рабочего времени по шаблону ```resources/template.docx```,
архивирует получившийся документ word и кидает ссылку Яндекс.Диск для скачивания. 

## Запуск бота
- Склонировать репозиторий
```bash
git clone https://github.com/Microsoft/botbuilder-python.git
```
- Заполнить переменные окружения в файле ```.env```
Пример файла
```.env
JIRA_SERVER=https://example.atlassian.net    #Адрес Jira 
JIRA_LOGIN=example@example.com
JIRA_TOKEN=JIRA_TOKEN
YA_TOKEN=YA_TOKEN
```
- Установить зависимости `pip install -r requirements.txt`
- Запустить `python app.py`

## Тестирование с помощью Bot Framework Emulator
[Microsoft Bot Framework Emulator](https://github.com/microsoft/botframework-emulator) десктопное приложение, которое позволяет протестировалть бота локально

- Скачать можно [здесь](https://github.com/Microsoft/BotFramework-Emulator/releases)

### Подключение бота к Bot Framework Emulator
- Запустите Bot Framework Emulator
- File -> Open Bot
- Укажите этот url - http://localhost:3978/api/messages

## Деплой бота в Azure

Подробнее о деплое бота можно почитать по ссылке [Deploy your bot to Azure](https://aka.ms/azuredeployment)
