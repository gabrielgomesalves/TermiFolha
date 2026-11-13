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
        self.run(self.m.criar_folha_dummy)


@pytest.fixture
def app(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(app_module, "limpar", lambda: None)
    monkeypatch.setattr(app_module, "pausar", lambda: None)
    monkeypatch.setattr(app_module, "FICHEIRO_ESTADO", str(tmp_path / "estado.pkl"))
    monkeypatch.setattr(app_module, "tabela", app_module.pd.DataFrame())
    monkeypatch.setattr(app_module, "tipos", {})
    app_module.undo.clear()
    app_module.redo.clear()
    app_module.undo.append((app_module.tabela.copy(), {}))
    yield App(app_module, monkeypatch)
    app_module.undo.clear()
    app_module.redo.clear()
