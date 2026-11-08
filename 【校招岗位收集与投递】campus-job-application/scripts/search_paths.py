"""One external-data location contract for the search component's public entries."""
import os
from pathlib import Path

def configure(workspace=None):
    value=workspace or os.environ.get('CAMPUS_RECRUITING_WORKSPACE')
    if not value:raise ValueError('CAMPUS_RECRUITING_WORKSPACE_REQUIRED')
    root=Path(value).resolve()
    private=Path(os.environ.get('CAMPUS_SEARCH_PRIVATE_DIR',root/'.runtime'/'search')).resolve()
    package=Path(__file__).resolve().parents[1]
    if root.is_relative_to(package) or private.is_relative_to(package):
        raise ValueError('USER_DATA_OUTSIDE_COMPONENT_REQUIRED')
    os.environ['CAMPUS_RECRUITING_WORKSPACE']=str(root)
    os.environ['CAMPUS_SEARCH_PRIVATE_DIR']=str(private)
    return root,private
