# setup_passwords.py
# Ejecuta este script para agregar o cambiar usuarios y contraseñas.

import yaml
from yaml.loader import SafeLoader
import bcrypt

CREDENTIALS_FILE = "credentials.yaml"

def hash_password(plain: str) -> str:
    return bcrypt.hashpw(plain.encode(), bcrypt.gensalt()).decode()

print("=" * 60)
print("  CONFIGURACIÓN DE USUARIOS — Materiales en Tránsito")
print("=" * 60)
print()

with open(CREDENTIALS_FILE, encoding="utf-8") as f:
    config = yaml.load(f, Loader=SafeLoader)

while True:
    print("Opciones:")
    print("  1. Agregar nuevo usuario")
    print("  2. Cambiar contraseña de usuario existente")
    print("  3. Eliminar usuario")
    print("  4. Ver usuarios actuales")
    print("  5. Salir")
    opcion = input("\nElige una opción (1-5): ").strip()

    if opcion == "1":
        username = input("  Nombre de usuario (sin espacios): ").strip().lower()
        if username in config["credentials"]["usernames"]:
            print("  Ese usuario ya existe.")
            continue
        name  = input("  Nombre completo: ").strip()
        email = input("  Email: ").strip()
        while True:
            pwd = input("  Contraseña (mínimo 6 caracteres): ").strip()
            if len(pwd) >= 6:
                break
            print("  Mínimo 6 caracteres.")
        config["credentials"]["usernames"][username] = {
            "name":     name,
            "email":    email,
            "password": hash_password(pwd),
        }
        print(f"  ✓ Usuario '{username}' agregado.")

    elif opcion == "2":
        users = list(config["credentials"]["usernames"].keys())
        print("  Usuarios:", ", ".join(users))
        username = input("  ¿Cuál usuario?: ").strip().lower()
        if username not in config["credentials"]["usernames"]:
            print("  Usuario no encontrado.")
            continue
        while True:
            pwd = input("  Nueva contraseña (mínimo 6 caracteres): ").strip()
            if len(pwd) >= 6:
                break
            print("  Mínimo 6 caracteres.")
        config["credentials"]["usernames"][username]["password"] = hash_password(pwd)
        print(f"  ✓ Contraseña de '{username}' actualizada.")

    elif opcion == "3":
        users = list(config["credentials"]["usernames"].keys())
        print("  Usuarios:", ", ".join(users))
        username = input("  ¿Cuál usuario eliminar?: ").strip().lower()
        if username not in config["credentials"]["usernames"]:
            print("  Usuario no encontrado.")
            continue
        confirm = input(f"  ¿Seguro que quieres eliminar '{username}'? (s/n): ").strip().lower()
        if confirm == "s":
            del config["credentials"]["usernames"][username]
            print(f"  ✓ Usuario '{username}' eliminado.")

    elif opcion == "4":
        print()
        for u, data in config["credentials"]["usernames"].items():
            print(f"  {u:<20} {data.get('name',''):<25} {data.get('email','')}")
        print()

    elif opcion == "5":
        break

with open(CREDENTIALS_FILE, "w", encoding="utf-8") as f:
    yaml.dump(config, f, allow_unicode=True, default_flow_style=False)

print("\n  ✓ Cambios guardados en credentials.yaml")
input("  Presiona Enter para salir...")
