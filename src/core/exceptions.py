"""Custom exceptions and standardized error catalog for Word operations."""

from typing import Dict


class LiveWordError(Exception):
    """Base exception for all Live Word COM interactions with user-friendly messages."""
    pass


ERROR_MESSAGES: Dict[str, str] = {
    "index_out_of_range": (
        "Index [p:{target_index}] is out of range ({valid_range})."
    ),
    "read_only_or_protected": (
        "The document '{file_name}' is read-only or protected ({state}). "
        "Please enable editing in Microsoft Word."
    ),
    "protected_view": (
        "The document '{file_name}' is open in Protected View. "
        "Please click 'Enable Editing' in Microsoft Word."
    ),
    "cloud_hosted": (
        "The document '{file_name}' is cloud-hosted on OneDrive/SharePoint ({url}). "
        "Please save or copy it locally, or close it in Word and set WORD_BACKEND=headless."
    ),
    "modal_dialog_or_timeout": (
        "Word may be waiting on a modal dialog or is not responding "
        "(call timed out after {timeout}s). Please check Microsoft Word and dismiss any open dialogs."
    ),
    "worker_poisoned": (
        "An earlier Word operation has not returned; Word is probably still blocked by a modal dialog. "
        "Please check Microsoft Word and dismiss any open dialogs."
    ),
    "not_open_in_word": (
        "WORD_BACKEND is set to 'live', but the document '{file_name}' is not currently open in Microsoft Word."
    ),
    "lock_file_cannot_attach": (
        "The document '{file_name}' appears to be open in Word (lock file found), "
        "but could not be attached via COM. Word may be running elevated (as Administrator), "
        "in a different session, or the lock file is stale. Please close Word or verify the lock."
    ),
    "tool_not_yet_live_capable": (
        "Tool '{tool_name}' is not yet live-capable in this phase. "
        "The file is open in Microsoft Word, so no changes were written to disk."
    ),
}
