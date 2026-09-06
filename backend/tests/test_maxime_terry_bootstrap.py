from pathlib import Path

from app.security import verify_password


def test_maxime_terry_account_is_active_and_uses_requested_password() -> None:
    migration = (
        Path(__file__).resolve().parents[1] / "migrations" / "007_bootstrap_maxime_terry.sql"
    ).read_text(encoding="utf-8")

    password_hash = next(part for part in migration.split("'") if part.startswith("$argon2id$"))
    assert "username_normalized = 'maximeterry'" in migration
    assert "'Maxime Terry'" in migration
    assert "'teacher'" in migration
    assert "'active'" in migration
    assert "must_change_password = false" in migration
    assert "'mty'" not in migration
    assert verify_password(password_hash, "mty")
