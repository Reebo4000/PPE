import os
from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_user, login_required, logout_user, current_user
from werkzeug.utils import secure_filename

from . import db
from .forms import LoginForm, RegistrationForm, UploadForm
from .models import User
from .utils import detect_helmet, save_violation, capture_image_from_ip

bp = Blueprint('main', __name__)


@bp.route('/')
@login_required
def index():
    return render_template('index.html')


@bp.route('/login', methods=['GET', 'POST'])
def login():
    form = LoginForm()
    if form.validate_on_submit():
        user = User.query.filter_by(username=form.username.data).first()
        if user and user.check_password(form.password.data):
            login_user(user)
            return redirect(url_for('main.index'))
        flash('Invalid credentials')
    return render_template('login.html', form=form)


@bp.route('/register', methods=['GET', 'POST'])
def register():
    form = RegistrationForm()
    if form.validate_on_submit():
        if User.query.filter_by(username=form.username.data).first():
            flash('Username already exists')
        else:
            user = User(username=form.username.data, role=form.role.data)
            user.set_password(form.password.data)
            db.session.add(user)
            db.session.commit()
            flash('Registration successful')
            return redirect(url_for('main.login'))
    return render_template('register.html', form=form)


@bp.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('main.login'))


@bp.route('/upload', methods=['GET', 'POST'])
@login_required
def upload_image():
    form = UploadForm()
    if form.validate_on_submit():
        file = form.image.data
        filename = secure_filename(file.filename)
        path = os.path.join('static', filename)
        file.save(path)
        if not detect_helmet(path):
            save_violation(current_user, path)
            flash('Violation detected')
        else:
            flash('Helmet detected')
    return render_template('upload.html', form=form)


@bp.route('/violations')
@login_required
def violations():
    if current_user.role != 'admin':
        user_violations = current_user.violations
    else:
        user_violations = []
        for user in User.query.all():
            user_violations.extend(user.violations)
    return render_template('violations.html', violations=user_violations)


@bp.route('/capture')
@login_required
def capture():
    ip = request.args.get('ip')
    username = request.args.get('username')
    password = request.args.get('password')
    if not ip:
        flash('IP address required')
        return redirect(url_for('main.index'))
    path = capture_image_from_ip(ip, username, password)
    if path and not detect_helmet(path):
        save_violation(current_user, path)
        flash('Violation detected')
    elif path:
        flash('Helmet detected')
    else:
        flash('Failed to capture image')
    return redirect(url_for('main.index'))

