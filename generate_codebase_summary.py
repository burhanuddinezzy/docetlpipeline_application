import ast
import os
from pathlib import Path
import json
from typing import List, Dict

CODEBASE_ROOT = Path(".")
SUMMARY_FILE = Path("CODEBASE_SUMMARY.md")
INDEX_FILE = Path(".summary_index.json")
IGNORE_PATTERNS = [
    ".git",         # ignore git folder
    "__pycache__",  # ignore cache folders
    "tests",        # ignore test folders
    "setup.py",     # ignore specific files
    "venv",
    "README.md",
    "whiteboard.md",
    "generate_codebase_summary.py"
]

def should_ignore(file_path: Path) -> bool:
    for pattern in IGNORE_PATTERNS:
        if pattern in file_path.parts or file_path.match(pattern):
            return True
    return False

def get_file_mtime(file_path: Path) -> float:
    return file_path.stat().st_mtime

def read_index() -> Dict[str, float]:
    if INDEX_FILE.exists():
        with open(INDEX_FILE, "r") as f:
            return json.load(f)
    return {}

def write_index(index: Dict[str, float]):
    with open(INDEX_FILE, "w") as f:
        json.dump(index, f)

# --- AST ANALYSIS ---
class FileAnalyzer(ast.NodeVisitor):
    def __init__(self, file_path: Path):
        self.file_path = str(file_path)
        self.imports: List[str] = []
        self.classes: Dict[str, Dict] = {}
        self.functions: Dict[str, Dict] = {}
        self.top_level_calls: List[str] = []
        self.if_main_calls: List[str] = []

    def visit_Import(self, node):
        for alias in node.names:
            self.imports.append(f"import {alias.name}" + (f" as {alias.asname}" if alias.asname else ""))
        self.generic_visit(node)

    def visit_ImportFrom(self, node):
        module = node.module or ""
        for alias in node.names:
            self.imports.append(f"from {module} import {alias.name}" + (f" as {alias.asname}" if alias.asname else ""))
        self.generic_visit(node)

    def visit_ClassDef(self, node):
        cls_name = node.name
        cls_dict = {"methods": {}, "calls": [], "bases": [ast.unparse(b) for b in node.bases]}
        for n in node.body:
            if isinstance(n, ast.FunctionDef):
                method_name = n.name
                calls, loops = self._extract_calls_and_loops(n)
                cls_dict["methods"][method_name] = {
                    "calls": calls,
                    "loops": loops
                }
                cls_dict["calls"].extend(calls)
        self.classes[cls_name] = cls_dict
        self.generic_visit(node)

    def visit_FunctionDef(self, node):
        func_name = node.name
        calls, loops = self._extract_calls_and_loops(node)
        self.functions[func_name] = {
            "calls": calls,
            "loops": loops
        }
        self.generic_visit(node)

    def visit_If(self, node):
        if (
            isinstance(node.test, ast.Compare) and
            isinstance(node.test.left, ast.Name) and
            node.test.left.id == "__name__"
        ):
            if any(isinstance(c, ast.Constant) and c.value == "__main__" for c in node.test.comparators):
                calls, _ = self._extract_calls_and_loops(node)
                self.if_main_calls.extend(calls)
        self.generic_visit(node)

    def visit_Expr(self, node):
        calls, _ = self._extract_calls_and_loops(node)
        self.top_level_calls.extend(calls)
        self.generic_visit(node)

    def _extract_calls_and_loops(self, node):
        calls = []
        loops = []

        class Visitor(ast.NodeVisitor):
            def visit_Call(self, call_node):
                if isinstance(call_node.func, ast.Name):
                    calls.append(call_node.func.id)
                elif isinstance(call_node.func, ast.Attribute):
                    calls.append(ast.unparse(call_node.func))
                self.generic_visit(call_node)

            def visit_For(self, n):
                loops.append("for")
                self.generic_visit(n)
            def visit_While(self, n):
                loops.append("while")
                self.generic_visit(n)

        Visitor().visit(node)
        return calls, loops


# --- SUMMARIZE FILE ---
def summarize_file(file_path: Path) -> Dict:
    try:
        source = file_path.read_text(encoding="utf-8")
        tree = ast.parse(source)
    except Exception as e:
        return {"error": str(e)}

    analyzer = FileAnalyzer(file_path)
    analyzer.visit(tree)
    return {
        "file": str(file_path),
        "imports": analyzer.imports,
        "classes": analyzer.classes,
        "functions": analyzer.functions,
        "top_level_calls": analyzer.top_level_calls,
        "if_main_calls": analyzer.if_main_calls
    }

# --- BUILD CALL GRAPH ---
def build_call_graph(file_summaries: List[Dict]) -> Dict[str, List[str]]:
    graph = {}
    for summary in file_summaries:
        if "error" in summary:
            continue
        file_prefix = Path(summary["file"]).stem
        # functions
        for func_name, func_info in summary["functions"].items():
            caller = f"{file_prefix}.{func_name}"
            graph[caller] = [f"{file_prefix}.{c}" for c in func_info["calls"]]
        # methods
        for cls_name, cls_info in summary["classes"].items():
            for m_name, m_info in cls_info["methods"].items():
                caller = f"{file_prefix}.{cls_name}.{m_name}"
                graph[caller] = [f"{file_prefix}.{c}" for c in m_info["calls"]]
        # top-level calls
        graph[f"{file_prefix}.__top_level__"] = [f"{file_prefix}.{c}" for c in summary["top_level_calls"]]
        graph[f"{file_prefix}.__main__"] = [f"{file_prefix}.{c}" for c in summary["if_main_calls"]]
    return graph

# --- WRITE SUMMARY MARKDOWN ---
def write_summary(file_summaries: List[Dict], call_graph: Dict):
    with open(SUMMARY_FILE, "w", encoding="utf-8") as f:
        f.write("# CODEBASE SUMMARY\n\n")
        for summary in file_summaries:
            if "error" in summary:
                f.write(f"**ERROR parsing {summary['file']}: {summary['error']}**\n\n")
                continue
            f.write(f"## File: {summary['file']}\n")
            if summary["imports"]:
                f.write("### Imports\n")
                for imp in summary["imports"]:
                    f.write(f"- {imp}\n")
            if summary["classes"]:
                f.write("### Classes\n")
                for cls_name, cls_info in summary["classes"].items():
                    f.write(f"- Class `{cls_name}` (bases: {', '.join(cls_info['bases'])})\n")
                    for m_name, m_info in cls_info["methods"].items():
                        f.write(f"  - Method `{m_name}` calls: {', '.join(m_info['calls']) or 'None'}, loops: {', '.join(m_info['loops']) or 'None'}\n")
            if summary["functions"]:
                f.write("### Functions\n")
                for f_name, f_info in summary["functions"].items():
                    f.write(f"- Function `{f_name}` calls: {', '.join(f_info['calls']) or 'None'}, loops: {', '.join(f_info['loops']) or 'None'}\n")
            if summary["top_level_calls"]:
                f.write(f"### Top-level calls: {', '.join(summary['top_level_calls'])}\n")
            if summary["if_main_calls"]:
                f.write(f"### if __name__ == '__main__' calls: {', '.join(summary['if_main_calls'])}\n")
            f.write("\n---\n")

        # Write call graph at the end
        f.write("\n# Call Graph\n")
        for caller, callees in call_graph.items():
            f.write(f"- `{caller}` calls: {', '.join(callees) or 'None'}\n")

# --- MAIN ---
def main():
    print("Starting code summary generation...")
    index = read_index()
    new_index = {}
    summaries = []

    # Debug: Print all Python files found
    all_files = [f for f in CODEBASE_ROOT.rglob("*.py")]
    print(f"\nFound {len(all_files)} Python files:")
    for f in all_files:
        print(f"- {f}")

    # Debug: Print files after ignore filter
    files_after_ignore = [f for f in all_files if not should_ignore(f)]
    print(f"\nAfter ignore filter, {len(files_after_ignore)} files remain:")
    for f in files_after_ignore:
        print(f"- {f}")

    for file_path in files_after_ignore:
        mtime = get_file_mtime(file_path)
        new_index[str(file_path)] = mtime

        # Debug: Print file being processed
        print(f"\nProcessing: {file_path}")
        
        # simple incremental update: skip unchanged files
        if str(file_path) in index and index[str(file_path)] == mtime:
            print(f"Skipping unchanged file: {file_path}")
            continue

        summary = summarize_file(file_path)
        summaries.append(summary)
        print(f"Added summary for: {file_path}")

    if not summaries:
        print("\nNo files to process! Check that:")
        print("1. Python files exist in the directory")
        print("2. Files aren't all being ignored")
        print("3. Files have been modified since last run")
        return

    call_graph = build_call_graph(summaries)
    write_summary(summaries, call_graph)
    write_index(new_index)
    print(f"\nSummary with call graph written to {SUMMARY_FILE}")

if __name__ == "__main__":
    main()
