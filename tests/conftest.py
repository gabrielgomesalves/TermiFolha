"""Fixtures partilhadas: isolam o estado global do TermiFolha e simulam o teclado."""
import builtins

import matplotlib
import pytest

matplotlib.use("Agg")

import TermiFolha as app_module


class App:
    """Envolve o módulo TermiFolha com um simulador de entradas do utilizador."""

    def __init__(self, module, monkeypatch):
        self.m = module
        self._mp = monkeypatch

    def feed(self, *answers):
        """Define as respostas que `input()` devolve; esgotadas, levanta AssertionError."""
        it = iter(answers)

        def fake_input(prompt=""):
            try:
                return next(it)
            except StopIteration:
                raise AssertionError(f"input() inesperado: {prompt!r}") from None

        self._mp.setattr(builtins, "input", fake_input)

    def run(self, func, *answers):
        self.feed(*answers)
        return func()

    def dummy(self):
        self.run(self.m.create_dummy_sheet)


@pytest.fixture
def app(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(app_module, "clear_screen", lambda: None)
    monkeypatch.setattr(app_module, "pause", lambda: None)
    monkeypatch.setattr(app_module, "STATE_FILE", str(tmp_path / "estado.json"))
    monkeypatch.setattr(app_module, "LEGACY_STATE_FILE", str(tmp_path / "estado.pkl"))
    monkeypatch.setattr(app_module, "table", app_module.pd.DataFrame())
    monkeypatch.setattr(app_module, "col_types", {})
    monkeypatch.setattr(app_module, "meta", app_module.new_meta())
    monkeypatch.setattr(app_module, "sheets", {app_module.DEFAULT_SHEET_NAME: None})
    monkeypatch.setattr(app_module, "active_name", app_module.DEFAULT_SHEET_NAME)
    monkeypatch.setattr(app_module, "settings", {"history_limit": app_module.HISTORY_LIMIT, "page_size": app_module.PAGE_SIZE})
    monkeypatch.setattr(app_module, "plot_custom", {"title": "", "xlabel": "", "ylabel": ""})
    app_module.undo_stack.clear()
    app_module.redo_stack.clear()
    app_module.undo_stack.append(app_module.snapshot())
    yield App(app_module, monkeypatch)
    app_module.undo_stack.clear()
    app_module.redo_stack.clear()
