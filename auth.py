"""Authentication and per-page permissions for the OIA Hub.

Each "controlled" page (Universities, Masters Data, Student Exchange,
Internships, Reports) can be edited only by users the admin has granted
control of that specific page — everyone signed in can still *view* every
page, they just can't make changes to pages they don't control. The admin
account can control every page and manages who else controls what from the
Admin page.

Users are stored in a small JSON file (data/users.json) rather than a real
database, since the rest of this app already reads its data from Excel
files on disk rather than a DB. Passwords are hashed with werkzeug's
generate_password_hash/check_password_hash — never stored in plain text.
"""

import json
import os
import sys
import time
import threading
import secrets
from werkzeug.security import generate_password_hash, check_password_hash

USERS_FILE = os.path.join(os.path.dirname(__file__), "data", "users.json")

# The controllable pages. Keys are used internally (permission checks, JSON);
# labels are what's shown in the Admin UI and nav.
PAGE_KEYS = ["universities", "masters", "student_exchange", "internships", "reports"]
PAGE_LABELS = {
    "universities": "Universities",
    "masters": "Masters Data",
    "student_exchange": "Student Exchange",
    "internships": "Internships",
    "reports": "Reports",
}

DEFAULT_ADMIN_USERNAME = "admin"
# First-run password for the built-in admin account. Override it with the
# OIA_HUB_ADMIN_PASSWORD environment variable, and change it from the Admin
# page after first login (the Admin page shows a warning until you do).
DEFAULT_ADMIN_PASSWORD = "admin123"

MIN_PASSWORD_LENGTH = 8

# Guards read-modify-write cycles on users.json (the dev server is threaded).
_store_lock = threading.RLock()


def _initial_admin_password():
    return os.environ.get("OIA_HUB_ADMIN_PASSWORD") or DEFAULT_ADMIN_PASSWORD


def _default_store():
    return {
        "next_id": 2,
        "users": [
            {
                "id": 1,
                "username": DEFAULT_ADMIN_USERNAME,
                "password_hash": generate_password_hash(_initial_admin_password()),
                "is_admin": True,
                "controls": PAGE_KEYS.copy(),
            }
        ],
    }


def _validate_store(store):
    if not isinstance(store, dict) or not isinstance(store.get("users"), list):
        raise ValueError("users.json has an unexpected structure")
    for u in store["users"]:
        if not isinstance(u, dict) or "id" not in u or "username" not in u or "password_hash" not in u:
            raise ValueError("users.json contains a malformed user record")
    store.setdefault("next_id", max([u["id"] for u in store["users"]] + [0]) + 1)
    return store


def _load_store():
    with _store_lock:
        if not os.path.isfile(USERS_FILE):
            store = _default_store()
            _save_store(store)
            return store
        try:
            with open(USERS_FILE, "r", encoding="utf-8") as f:
                return _validate_store(json.load(f))
        except (json.JSONDecodeError, OSError, ValueError) as exc:
            # Never silently wipe every account: keep the unreadable file for
            # inspection, then start over with a fresh admin account.
            backup = f"{USERS_FILE}.corrupt-{int(time.time())}"
            try:
                os.replace(USERS_FILE, backup)
            except OSError:
                backup = "(could not be moved aside)"
            print(
                f"[auth] WARNING: {USERS_FILE} was unreadable ({exc}). "
                f"Saved as {backup} and re-created the default '{DEFAULT_ADMIN_USERNAME}' account.",
                file=sys.stderr,
            )
            store = _default_store()
            _save_store(store)
            return store


def _save_store(store):
    """Atomic write: a crash mid-save can't leave a half-written users.json."""
    with _store_lock:
        os.makedirs(os.path.dirname(USERS_FILE), exist_ok=True)
        tmp = f"{USERS_FILE}.tmp-{os.getpid()}-{threading.get_ident()}"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(store, f, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, USERS_FILE)
        try:
            os.chmod(USERS_FILE, 0o600)
        except OSError:
            pass


def list_users():
    """Users without their password hashes, for display."""
    store = _load_store()
    return [
        {k: v for k, v in u.items() if k != "password_hash"}
        for u in store["users"]
    ]


def find_by_username(username):
    store = _load_store()
    needle = (username or "").strip().lower()
    for u in store["users"]:
        if u["username"].lower() == needle:
            return u
    return None


def find_by_id(user_id):
    store = _load_store()
    for u in store["users"]:
        if u["id"] == user_id:
            return u
    return None


# A throwaway hash so "unknown user" and "wrong password" take the same time
# to reject, which stops the login form from revealing which usernames exist.
_DUMMY_HASH = generate_password_hash(secrets.token_hex(8))


def verify_login(username, password):
    user = find_by_username(username)
    if user is None:
        check_password_hash(_DUMMY_HASH, password or "")
        return None
    if check_password_hash(user["password_hash"], password or ""):
        return user
    return None


def can_control(user, page_key):
    """Whether this user may edit/manage the given page's data."""
    if not user:
        return False
    if user.get("is_admin"):
        return True
    return page_key in (user.get("controls") or [])


def default_admin_password_in_use():
    """True while the built-in admin account still has the shipped default password."""
    user = find_by_username(DEFAULT_ADMIN_USERNAME)
    return bool(
        user
        and user.get("is_admin")
        and check_password_hash(user["password_hash"], DEFAULT_ADMIN_PASSWORD)
    )


def _check_password_rules(password):
    if not password:
        raise ValueError("Password is required.")
    if len(password) < MIN_PASSWORD_LENGTH:
        raise ValueError(f"Password must be at least {MIN_PASSWORD_LENGTH} characters.")


def create_user(username, password, is_admin=False, controls=None):
    username = (username or "").strip()
    if not username:
        raise ValueError("Username is required.")
    _check_password_rules(password)

    with _store_lock:
        if find_by_username(username):
            raise ValueError(f'A user named "{username}" already exists.')
        store = _load_store()
        user = {
            "id": store["next_id"],
            "username": username,
            "password_hash": generate_password_hash(password),
            "is_admin": bool(is_admin),
            "controls": [c for c in (controls or []) if c in PAGE_KEYS],
        }
        store["next_id"] += 1
        store["users"].append(user)
        _save_store(store)
        return user


def update_permissions(user_id, is_admin, controls):
    with _store_lock:
        store = _load_store()
        target = next((u for u in store["users"] if u["id"] == user_id), None)
        if target is None:
            raise ValueError("User not found.")
        if target.get("is_admin") and not is_admin:
            others = [u for u in store["users"] if u["id"] != user_id and u.get("is_admin")]
            if not others:
                raise ValueError("Can't remove admin access from the last remaining admin account.")
        target["is_admin"] = bool(is_admin)
        target["controls"] = [c for c in (controls or []) if c in PAGE_KEYS]
        _save_store(store)
        return target


def update_password(user_id, new_password):
    _check_password_rules(new_password)
    with _store_lock:
        store = _load_store()
        for u in store["users"]:
            if u["id"] == user_id:
                u["password_hash"] = generate_password_hash(new_password)
                _save_store(store)
                return u
    raise ValueError("User not found.")


def delete_user(user_id):
    with _store_lock:
        store = _load_store()
        remaining = [u for u in store["users"] if u["id"] != user_id]
        if len(remaining) == len(store["users"]):
            raise ValueError("User not found.")
        if not any(u.get("is_admin") for u in remaining):
            raise ValueError("Can't delete the last remaining admin account.")
        store["users"] = remaining
        _save_store(store)


def reset_admin(username=None, password=None):
    """Recovery for a locked-out admin: (re)create or repair an admin account.

    Run from the project folder:  python auth.py reset-admin [username]
    Returns (username, password) so the caller can show the new password.
    """
    username = (username or DEFAULT_ADMIN_USERNAME).strip()
    password = password or secrets.token_urlsafe(12)
    _check_password_rules(password)
    with _store_lock:
        store = _load_store()
        needle = username.lower()
        user = next((u for u in store["users"] if u["username"].lower() == needle), None)
        if user:
            user["password_hash"] = generate_password_hash(password)
            user["is_admin"] = True
            user["controls"] = PAGE_KEYS.copy()
        else:
            store["users"].append({
                "id": store["next_id"],
                "username": username,
                "password_hash": generate_password_hash(password),
                "is_admin": True,
                "controls": PAGE_KEYS.copy(),
            })
            store["next_id"] += 1
        _save_store(store)
    return username, password


if __name__ == "__main__":
    if len(sys.argv) >= 2 and sys.argv[1] == "reset-admin":
        name, pw = reset_admin(sys.argv[2] if len(sys.argv) > 2 else None,
                               sys.argv[3] if len(sys.argv) > 3 else None)
        print(f"Admin account '{name}' is ready. Password: {pw}")
        print("Sign in and change this password from the Admin page.")
    else:
        print("Usage: python auth.py reset-admin [username] [new-password]")
