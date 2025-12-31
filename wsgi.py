# WSGI entrypoint for Gunicorn
from webtool.app import app as application

# Optional: expose as `app` too
app = application
