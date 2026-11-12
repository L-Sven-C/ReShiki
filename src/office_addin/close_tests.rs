use super::*;

const METADATA: &[u8] = br#"{"version":1,"sessionId":"00000000-0000-0000-0000-000000000001"}"#;

#[test]
fn tab_close_publishes_only_a_matching_notice_and_is_idempotent() {
    let directory = tempfile::tempdir().unwrap();
    let draft = directory.path().join("drawing.rsk");
    for (name, bytes) in [
        ("session.json", METADATA),
        ("drawing.rsk", b"latest draft".as_slice()),
        ("request.json", b"pending save".as_slice()),
    ] {
        std::fs::write(directory.path().join(name), bytes).unwrap();
    }
    finish(&draft).unwrap();
    finish(&draft).unwrap();
    let notice: serde_json::Value =
        serde_json::from_slice(&std::fs::read(directory.path().join("closed.json")).unwrap())
            .unwrap();
    assert_eq!(
        notice,
        serde_json::from_slice::<serde_json::Value>(METADATA).unwrap()
    );
    assert_eq!(
        std::fs::read(directory.path().join("session.json")).unwrap(),
        METADATA
    );
    assert_eq!(std::fs::read(&draft).unwrap(), b"latest draft");
    assert_eq!(
        std::fs::read(directory.path().join("request.json")).unwrap(),
        b"pending save"
    );
    assert!(!directory.path().join("ack.json").exists());
}

#[test]
fn tab_close_rejects_invalid_metadata_and_retains_the_draft_on_write_failure() {
    let directory = tempfile::tempdir().unwrap();
    let draft = directory.path().join("drawing.rsk");
    std::fs::write(&draft, b"latest draft").unwrap();
    for metadata in [b"{".as_slice(), b"{}", &[b' '; 4097]] {
        std::fs::write(directory.path().join("session.json"), metadata).unwrap();
        assert!(finish(&draft).is_err());
        assert!(!directory.path().join("closed.json").exists());
    }
    std::fs::write(directory.path().join("session.json"), METADATA).unwrap();
    std::fs::create_dir(directory.path().join("closed.json")).unwrap();
    assert!(finish(&draft).is_err());
    assert_eq!(std::fs::read(&draft).unwrap(), b"latest draft");
    assert_eq!(
        std::fs::read(directory.path().join("session.json")).unwrap(),
        METADATA
    );
    assert!(!directory.path().join("ack.json").exists());
}

#[test]
fn already_retired_session_needs_no_tab_close_notice() {
    let directory = tempfile::tempdir().unwrap();
    std::fs::write(
        directory.path().join("session.json"),
        br#"{"version":1,"sessionId":"00000000-0000-0000-0000-000000000001","closed":true}"#,
    )
    .unwrap();
    finish(&directory.path().join("drawing.rsk")).unwrap();
    assert!(!directory.path().join("closed.json").exists());
}

#[cfg(unix)]
#[test]
fn tab_close_rejects_symlinked_session_metadata() {
    let directory = tempfile::tempdir().unwrap();
    let outside = tempfile::NamedTempFile::new().unwrap();
    std::fs::write(outside.path(), METADATA).unwrap();
    std::os::unix::fs::symlink(outside.path(), directory.path().join("session.json")).unwrap();
    assert!(finish(&directory.path().join("drawing.rsk")).is_err());
    assert_eq!(std::fs::read(outside.path()).unwrap(), METADATA);
    assert!(!directory.path().join("closed.json").exists());
}
