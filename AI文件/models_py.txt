# modules/models.py
from flask_login import UserMixin

class User(UserMixin):
    def __init__(self, id, username=None):
        self.id = str(id)
        self.username = username