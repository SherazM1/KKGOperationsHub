"""Shared targeting rules for the optional multistop comment."""
COMMENT_SCOPES = ('Master only', 'All BOLs', 'Children only')


def include_bol_comment(enabled: bool, scope: str, *, is_master: bool) -> bool:
    if scope not in COMMENT_SCOPES:
        raise ValueError(f'Unknown BOL comment scope: {scope}')
    return bool(enabled and (scope == 'All BOLs' or
                            (scope == 'Master only' and is_master) or
                            (scope == 'Children only' and not is_master)))
