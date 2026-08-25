# app/routes/__init__.py
from app.routes import auth, main, upload, labels, materials, admin, settings_routes

def register_routes(app):
    """Регистрирует все маршруты"""
    app.register_blueprint(auth.auth_bp)
    app.register_blueprint(main.main_bp)
    app.register_blueprint(upload.upload_bp)
    app.register_blueprint(labels.labels_bp)
    app.register_blueprint(materials.materials_bp)
    app.register_blueprint(admin.admin_bp)
    app.register_blueprint(settings_routes.settings_bp)