"""
excel_tool/registry.py — Action handler registry and project registration.

Provides the plumbing that connects action keywords (TOUCH, WAIT_FOR, …)
to their generator-method implementations:

- ``HandlerRegistry[T]`` — a generic registry mapping string keys to handlers
- ``action(name)`` — decorator to tag a generator method as an action handler
- ``_GenMeta`` — metaclass that collects @action()-tagged methods per class
- ``register_project(name)`` — decorator to register a project generator
- ``_PROJECT_REGISTRY`` — the global project-name → generator-class mapping
"""

from __future__ import annotations

from typing import Callable, Generic, TypeVar

T = TypeVar("T")


# ── Generic Handler Registry ──────────────────────────────────────────────────


class HandlerRegistry(Generic[T]):
    """A generic registry mapping string keys to handler objects.

    Usage::

        reg = HandlerRegistry()

        @reg.register("TOUCH")
        def handle_touch(step, ctx): ...

        handler = reg.get("TOUCH")   # → handle_touch
    """

    def __init__(self) -> None:
        self._handlers: dict[str, T] = {}

    def register(self, name: str) -> Callable[[T], T]:
        """Decorator: register *handler* under the given *name* (case-insensitive)."""

        def deco(handler: T) -> T:
            self._handlers[name.upper()] = handler
            return handler

        return deco

    def get(self, name: str, default: T | None = None) -> T | None:
        """Look up a handler by *name* (case-insensitive)."""
        return self._handlers.get(name.upper(), default)

    @property
    def all(self) -> dict[str, T]:
        """Return a shallow copy of every registered handler."""
        return dict(self._handlers)


# ── Legacy Action Decorator ───────────────────────────────────────────────────


def action(name: str) -> Callable:
    """Decorator: tag a method as the handler for an action keyword.

    Subclasses inherit parents' handlers via the ``_GenMeta`` metaclass;
    a child decorating the same action name overrides its parent's handler.

    Usage::

        @action("TOUCH")
        def handle_touch(self, step, ctx): ...
    """

    def deco(fn: Callable) -> Callable:
        fn._action_name = name.upper()  # type: ignore[attr-defined]
        return fn

    return deco


# ── Metaclass (collects @action methods) ──────────────────────────────────────


class _GenMeta(type):
    """Metaclass that collects every ``@action()``-tagged method into a per-class
    ``_HANDLERS`` dict.

    Subclasses inherit parents' handlers first; their own ``@action``-tagged
    methods override by action name.
    """

    def __new__(mcs, name: str, bases: tuple[type, ...], ns: dict) -> type:
        handlers: dict[str, Callable] = {}
        for base in reversed(bases):
            handlers.update(getattr(base, "_HANDLERS", {}))
        for v in ns.values():
            if callable(v) and hasattr(v, "_action_name"):
                handlers[v._action_name] = v  # type: ignore[arg-type]
        ns["_HANDLERS"] = handlers
        return super().__new__(mcs, name, bases, ns)


# ── Project Registry ──────────────────────────────────────────────────────────


_PROJECT_REGISTRY: dict[str, type] = {}
"""Global mapping: normalised project name → generator *class*."""


def register_project(name: str) -> Callable:
    """Decorator: register a generator subclass under a project name.

    Usage::

        @register_project("mygame")
        class MyGameGenerator(AirtestGenerator):
            DEFAULT_APP_PACKAGE = "com.my.game"
    """

    def deco(cls: type) -> type:
        _PROJECT_REGISTRY[name.lower()] = cls
        return cls

    return deco
