// Fork regression coverage: retain the newest draft even when session metadata
// cannot authorize a Word update. These are native filesystem tests, not Office E2E.
use super::*;

#[test]
fn invalid_session_metadata_preserves_latest_draft_without_publishing_a_request() {
    for metadata in [
        None,
        Some(b"{".to_vec()),
        Some(br#"{"version":2,"sessionId":"00000000-0000-0000-0000-000000000001"}"#.to_vec()),
        Some(br#"{"version":1,"sessionId":"invalid"}"#.to_vec()),
        Some(vec![b' '; 4097]),
    ] {
        let directory = tempfile::tempdir().unwrap();
        let draft = directory.path().join("drawing.rsk");
        std::fs::write(&draft, b"previous drawing").unwrap();
        if let Some(metadata) = metadata {
            std::fs::write(directory.path().join("session.json"), metadata).unwrap();
        }
        let error = save_with_timeout(&draft, b"latest drawing", Duration::ZERO).unwrap_err();
        assert_eq!(std::fs::read(&draft).unwrap(), b"latest drawing");
        assert!(error.contains("Draft saved at"), "{error}");
        assert!(error.contains(&draft.display().to_string()), "{error}");
        assert!(!directory.path().join("request.json").exists());
        assert!(!directory.path().join("ack.json").exists());
    }
}

#[test]
fn failed_draft_write_does_not_publish_a_save_request_or_claim_recovery() {
    let directory = tempfile::tempdir().unwrap();
    std::fs::write(
        directory.path().join("session.json"),
        br#"{"version":1,"sessionId":"00000000-0000-0000-0000-000000000001"}"#,
    )
    .unwrap();
    // A directory at the destination deterministically fails atomic replacement,
    // unlike permissions tests which can pass when run by a privileged user.
    let draft = directory.path().join("drawing.rsk");
    std::fs::create_dir(&draft).unwrap();
    let error = save_with_timeout(&draft, b"latest drawing", Duration::ZERO).unwrap_err();
    assert!(draft.is_dir());
    assert!(!error.contains("Draft saved at"));
    assert!(!directory.path().join("request.json").exists());
}

#[cfg(unix)]
#[test]
fn symlinked_session_is_rejected_but_the_native_draft_survives() {
    let directory = tempfile::tempdir().unwrap();
    let outside = tempfile::NamedTempFile::new().unwrap();
    let metadata = br#"{"version":1,"sessionId":"00000000-0000-0000-0000-000000000001"}"#;
    std::fs::write(outside.path(), metadata).unwrap();
    std::os::unix::fs::symlink(outside.path(), directory.path().join("session.json")).unwrap();
    let draft = directory.path().join("drawing.rsk");
    assert!(save_with_timeout(&draft, b"latest drawing", Duration::ZERO).is_err());
    assert_eq!(std::fs::read(draft).unwrap(), b"latest drawing");
    assert_eq!(std::fs::read(outside.path()).unwrap(), metadata);
    assert!(!directory.path().join("request.json").exists());
}
