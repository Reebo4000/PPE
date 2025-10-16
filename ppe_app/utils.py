import os
from datetime import datetime

import cv2

from .models import Violation, db


# Placeholder detection function
# Replace with actual model inference

def detect_helmet(image_path):
    # TODO: integrate actual helmet detection model
    # Here we simply return False to simulate violation
    return False


def save_violation(user, image_path, description='No helmet detected'):
    violation = Violation(image_path=image_path, description=description, user=user)
    db.session.add(violation)
    db.session.commit()


def capture_image_from_ip(ip, username=None, password=None):
    # This is a placeholder implementation that reads a frame from an IP camera URL
    # Actual integration will vary depending on camera make/model
    stream_url = f'http://{ip}/video'
    cap = cv2.VideoCapture(stream_url)
    ret, frame = cap.read()
    if ret:
        ts = datetime.utcnow().strftime('%Y%m%d_%H%M%S')
        path = os.path.join('static', f'capture_{ts}.jpg')
        cv2.imwrite(path, frame)
        cap.release()
        return path
    cap.release()
    return None
