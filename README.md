# EnvioMails

Tool to send emails with a defined text and attachments to a list of receivers defined in a Excel file.

## Configuration

To configure the application it is necessary create a configuration file with the following structure:

```
[EMAIL_CONF]
EMAIL_SERVER = {EMAIL_SERVER}
EMAIL_PORT = {EMAIL_PORT}
EMAIL_FROM = {EMAIL_ADDRESS}
EMAIL_PASSWD = {EMAIL_PASSWORD}

[EMAIL_MSG]
EMAIL_SUBJECT = {EMAIL_SUBJECT}
EMAIL_TEXT = {ROUTE_TO_FILE_WITH_TEXT}
EMAIL_ATTACHMENT = {ROUTE_TO_FILE_TO_ATTACH}
```

## Execution

To execute the application it is necessary launch the following command:

```
python3 EnvioCVs.py -c config.ini -d emails_to_send.xlsx
```

