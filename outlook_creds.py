import logging
import winreg

logging.basicConfig(format = '[%(asctime)s | %(levelname)s]: %(message)s', datefmt='%m.%d.%Y %H:%M:%S', level=logging.INFO)
log = logging.getLogger(__name__)

outlook_keys = (u'Software\\Microsoft\\Office\\15.0\\Outlook\\Profiles\\Outlook\\9375CFF0413111d3B88A00104B2A6676',
                u'Software\\Microsoft\\Office\\16.0\\Outlook\\Profiles\\Outlook\\9375CFF0413111d3B88A00104B2A6676',
                u'Software\\Microsoft\\Windows Messaging Subsystem\Profiles\\9375CFF0413111d3B88A00104B2A6676',
                u'Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows Messaging Subsystem\\Profiles\\Outlook\\9375CFF0413111d3B88A00104B2A6676')

def write_outlook_creds(key, email, passwd):
    try:
        key_handle = winreg.OpenKeyEx(reg, key, 0, winreg.KEY_ALL_ACCESS)
        log.info('Opened key: {}'.format(key))

    except Exception:
        key_handle = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_ALL_ACCESS)
        log.info('Created key: {}'.format(key))

    accountID = 1
    try:
        for i in range(winreg.QueryInfoKey(key_handle)[1]):
            vname, value, _ = winreg.EnumValue(key_handle, i)

            if vname != 'NextAccountID':
                continue

            accountId = value
            newAccountId = accountId + 1
            winreg.SetValueEx(key_handle, 'NextAccountID', 0, winreg.REG_DWORD, newAccountId)
            log.info('Writed new NextAccountID {} to {}'.format(newAccountId, key))

    except Exception:
        log.error('Failed reading NextAccountId')

    profile_key = key + '\\{0:08d}.format(accountID)'
    profile_key_handle = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, profile_key, 0, winreg.KEY_ALL_ACCESS)
    log.info('Created new profile: {}'.format(profile_key))

    try:
        winreg.SetValueEx(profile_key_handle, 'Email', 0, _winreg.REG_SZ, email)
        log.info('Writed email {} to {}'.format(email, key))

        winreg.SetValueEx(key_handle, 'IMAP Password', 0, winreg.REG_BINARY, password_blob)
        winreg.SetValueEx(key_handle, 'IMAP Password2', 0, winreg.REG_BINARY, password_blob)
        log.info('Writed IMAP password {} to {}'.format(passwd, key))

        winreg.SetValueEx(key_handle, 'POP3 Password', 0, winreg.REG_BINARY, password_blob)
        winreg.SetValueEx(key_handle, 'POP3 Password2',  0, winreg.REG_BINARY, password_blob)
        log.info('Writed POP3 password {} to {}'.format(passwd, key))

        winreg.SetValueEx(key_handle, 'HTTP Password', 0, winreg.REG_BINARY, password_blob)
        winreg.SetValueEx(key_handle, 'HTTP Password2', 0, winreg.REG_BINARY, password_blob)
        log.info('Writed HTTP password {} to {}'.format(passwd, key))

        winreg.SetValueEx(key_handle, 'SMTP Password', 0, winreg.REG_BINARY, password_blob)
        winreg.SetValueEx(key_handle, 'SMTP Password2', 0, winreg.REG_BINARY, password_blob)
        log.info('Writed SMTP password {} to {}'.format(passwd, key))

        winreg.SetValueEx(key_handle, 'HTTPMail Password', 0, winreg.REG_BINARY, password_blob)
        winreg.SetValueEx(key_handle, 'HTTPMail Password2', 0, winreg.REG_BINARY, password_blob)
        log.info('Writed HTTPMail password {} to {}'.format(passwd, key))

        winreg.SetValueEx(key_handle, 'NNTP Password', 0, winreg.REG_BINARY, password_blob)
        winreg.SetValueEx(key_handle, 'NNTP Password2', 0, _winreg.REG_BINARY, password_blob)
        log.info('Writed NNTP password {} to {}'.format(passwd, key))

    except Exception:
        log.error('Failed writing creds')

def write_outlook_any_creds(email, passwd):
    for key in outlook_keys:
        write_outlook_creds(key, email, passwd)


if __name__ == '__main__':
    email = 'test@mail.com'
    passwd = 'p@s$w0rd'

    write_outlook_any_creds(email, passwd)
