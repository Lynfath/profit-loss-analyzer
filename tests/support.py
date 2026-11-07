"""Load production helpers without executing Streamlit's top-level UI."""
import ast
from pathlib import Path
from types import SimpleNamespace


class State(dict):
    __getattr__ = dict.__getitem__
    __setattr__ = dict.__setitem__


APP = Path(__file__).resolve().parents[1] / "app.py"


def load_helpers(path=APP):
    tree = ast.parse(Path(path).read_text())
    nodes = [n for n in tree.body if isinstance(n, (ast.Import, ast.ImportFrom, ast.FunctionDef))
             or (isinstance(n, ast.Assign) and any(isinstance(t, ast.Name) and t.id == "PRICING_BASES" for t in n.targets))]
    ns = {}
    exec(compile(ast.Module(body=nodes, type_ignores=[]), str(path), "exec"), ns)
    ns["st"] = SimpleNamespace(session_state=State())
    ns["init_quote_state"]()
    return SimpleNamespace(**ns)
