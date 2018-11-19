# Python 3

from apiclient import errors

def print_about(service):
  """Print information about the user along with the Drive API settings.

  Args:
    service: Drive API service instance.
  """
  try:
    about = service.about().get().execute()

    print ('Current user name: %s' % about['name'])
    print ('Root folder ID: %s' % about['rootFolderId'])
    print ('Total quota (bytes): %s' % about['quotaBytesTotal'])
    print ('Used quota (bytes): %s' % about['quotaBytesUsed'])
  except errors.HttpError as error:
    print ('An error occurred: {}'.format(error))

print_about("Drive#about")