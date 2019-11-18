from O365 import Account, FileSystemTokenBackend
credentials = ('dc5c5f85-b8ff-43fd-b964-6c145fd1cae0', 'yEjGet8pglE/hY:gS/OpFL2oeg4=v81=')

account = Account(credentials)
if account.authenticate(scopes=['basic', 'message_all']):
   print('Authenticated!')