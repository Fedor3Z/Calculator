from __future__ import annotations

from collections import defaultdict, deque
from typing import Dict, Iterable, List, Set

from .formula_parser import extract_dependencies
from .model import normalize_cell


def build_dependency_graph(formulas: Dict[str, str]) -> Dict[str, Set[str]]:
    graph: Dict[str, Set[str]] = {}
    for cell, formula in formulas.items():
        graph[normalize_cell(cell)] = {normalize_cell(dep) for dep in extract_dependencies(formula)}
    return graph


def topological_sort(graph: Dict[str, Set[str]]) -> List[str]:
    indegree: Dict[str, int] = defaultdict(int)
    for node, deps in graph.items():
        indegree.setdefault(node, 0)
        for dep in deps:
            if dep in graph:
                indegree[node] += 1
    queue = deque([node for node, deg in indegree.items() if deg == 0])
    ordered: List[str] = []
    while queue:
        node = queue.popleft()
        ordered.append(node)
        for target, deps in graph.items():
            if node in deps and target in indegree:
                indegree[target] -= 1
                if indegree[target] == 0:
                    queue.append(target)
    if len(ordered) != len(graph):
        missing = set(graph) - set(ordered)
        raise ValueError(f"Циклические зависимости: {sorted(missing)}")
    return ordered


def dependency_chain(graph: Dict[str, Set[str]], cell: str) -> List[str]:
    cell = normalize_cell(cell)
    visited: Set[str] = set()
    order: List[str] = []

    def visit(node: str) -> None:
        for dep in graph.get(node, set()):
            if dep not in visited:
                visited.add(dep)
                order.append(dep)
                visit(dep)

    visit(cell)
    return order
