# This is a sample Python script.

from exchangelib import Account, Configuration, \
    OAuth2Credentials, Identity, OAUTH2
from exchangelib import HTMLBody
from exchangelib import Message, Mailbox


def mailsend(get_subject, htmlbody, get_to_recipients=[], get_bcc_recipients=[], get_cc_recipients=[]):
    application_id = "xxx"
    client_secret = "xxx"
    tenant_id = "xxx"
    address = "sender@mail.com"

    account = Account(
        primary_smtp_address=address,
        config=Configuration(
            credentials=OAuth2Credentials(
                client_id=application_id,
                client_secret=client_secret,
                tenant_id=tenant_id,
                identity=Identity(smtp_address=address)
            ),
            auth_type=OAUTH2,
            server="outlook.office365.com",
        ),
        autodiscover=False
    )

    Message(
        account=account,
        folder=account.sent,
        subject=get_subject,
        body=HTMLBody(
            '<html><body>{}</body></html>'.format(htmlbody),
        ),
        to_recipients=get_to_recipients,
        bcc_recipients=get_bcc_recipients,
        cc_recipients=get_cc_recipients,
    ).send_and_save()


if __name__ == '__main__':
    mailsend('PyCharm', 'Mailbody', ['mail@example.com'], ['mail@example.com'], ['mail@example.com'])
