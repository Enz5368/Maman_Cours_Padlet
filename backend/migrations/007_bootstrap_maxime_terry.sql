DO $$
DECLARE
    account_user_id uuid;
BEGIN
    SELECT id INTO account_user_id FROM users WHERE username_normalized = 'maximeterry';
    IF account_user_id IS NULL THEN
        account_user_id := gen_random_uuid();
        INSERT INTO users (
            id,
            username,
            username_normalized,
            password_hash,
            role,
            status,
            must_change_password,
            storage_quota_bytes,
            storage_used_bytes,
            display_name
        )
        VALUES (
            account_user_id,
            'Maxime Terry',
            'maximeterry',
            '$argon2id$v=19$m=65536,t=3,p=2$VdwImxsgc4KRqRPjsYSVTA$MbFsFetODnkFF0o5DtYYvxwYPKAQ0c5CycdMqHo2iZY',
            'teacher',
            'active',
            false,
            10737418240,
            0,
            'Maxime Terry'
        );
        INSERT INTO user_settings (user_id, settings_json) VALUES (account_user_id, '{}'::jsonb);
        INSERT INTO user_quotas (user_id, max_bytes, max_file_bytes) VALUES (account_user_id, 10737418240, 536870912);
        INSERT INTO user_workspaces (user_id, schema_version, revision, content) VALUES (account_user_id, 2, 1, '{}'::jsonb);
    ELSE
        UPDATE users
        SET username = 'Maxime Terry',
            display_name = 'Maxime Terry',
            password_hash = '$argon2id$v=19$m=65536,t=3,p=2$VdwImxsgc4KRqRPjsYSVTA$MbFsFetODnkFF0o5DtYYvxwYPKAQ0c5CycdMqHo2iZY',
            role = 'teacher',
            status = 'active',
            must_change_password = false
        WHERE id = account_user_id;
    END IF;
END $$;
