import win32com.client as win32

outlook = win32.Dispatch('outlook.application')

# see available mail accounts (no condition, just as info)
for accounts in outlook.Session.Accounts:
    print('Available email accounts: %s'%(accounts))