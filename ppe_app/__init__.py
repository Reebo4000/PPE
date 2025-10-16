from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager


# Initialize extensions
login_manager = LoginManager()
db = SQLAlchemy()


def create_app():
    app = Flask(__name__)
    app.config['SECRET_KEY'] = 'change-this-secret'
    app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///ppe.db'

    db.init_app(app)
    login_manager.init_app(app)
    login_manager.login_view = 'login'

    from . import routes  # noqa: W503
    app.register_blueprint(routes.bp)

    with app.app_context():
        db.create_all()

    return app
