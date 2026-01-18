from .model import (
    DocumentModel,
    Section,
    ElementType,
    ParagraphElement,
    TableElement,
    FigureElement,
    FormulaElement,
)


def print_document_structure(
    model: DocumentModel, show_paragraphs: bool = False
) -> str:
    lines = []
    lines.append("=" * 60)
    lines.append("СТРУКТУРА ДОКУМЕНТА")
    lines.append("=" * 60)

    lines.append(f"\nВсего элементов: {len(model.elements)}")
    lines.append(f"Заголовков: {len(model.headings)}")
    lines.append(f"Параграфов: {len(model.paragraphs)}")
    lines.append(f"Таблиц: {len(model.tables)}")
    lines.append(f"Рисунков: {len(model.figures)}")
    lines.append(f"Формул: {len(model.formulas)}")
    lines.append(f"Секций верхнего уровня: {len(model.sections)}")

    lines.append("\n" + "-" * 60)
    lines.append("ИЕРАРХИЯ СЕКЦИЙ")
    lines.append("-" * 60 + "\n")

    for section in model.sections:
        lines.extend(_print_section(section, 0, show_paragraphs))

    if model.tables:
        lines.append("\n" + "-" * 60)
        lines.append("ТАБЛИЦЫ")
        lines.append("-" * 60 + "\n")
        for table in model.tables:
            refs = (
                ", ".join(str(r) for r in table.referenced_at)
                if table.referenced_at
                else "нет"
            )
            status = "✓" if table.has_reference_before else "✗"
            lines.append(
                f"  [{table.index}] Таблица {table.number or '?'}: {table.caption or 'без подписи'}"
            )
            lines.append(f"       Ссылки: {refs} | До таблицы: {status}")

    if model.figures:
        lines.append("\n" + "-" * 60)
        lines.append("РИСУНКИ")
        lines.append("-" * 60 + "\n")
        for figure in model.figures:
            refs = (
                ", ".join(str(r) for r in figure.referenced_at)
                if figure.referenced_at
                else "нет"
            )
            status = "✓" if figure.has_reference_before else "✗"
            lines.append(
                f"  [{figure.index}] Рисунок {figure.number or '?'}: {figure.caption or 'без подписи'}"
            )
            lines.append(f"       Ссылки: {refs} | До рисунка: {status}")

    return "\n".join(lines)


def _print_section(section: Section, depth: int, show_paragraphs: bool) -> list[str]:
    lines = []
    indent = "  " * depth

    marker = "📁" if section.children else "📄"
    struct = " [структурный]" if section.heading.is_structural else ""
    num = f"({section.number}) " if section.number else ""

    lines.append(f"{indent}{marker} {num}{section.title[:60]}{struct}")

    if show_paragraphs and section.elements:
        for elem in section.elements:
            elem_indent = "  " * (depth + 1)
            if isinstance(elem, ParagraphElement):
                text = elem.text[:50] + "..." if len(elem.text) > 50 else elem.text
                lines.append(f"{elem_indent}  ¶ {text}")
            elif isinstance(elem, TableElement):
                lines.append(f"{elem_indent}  📊 Таблица {elem.number or '?'}")
            elif isinstance(elem, FigureElement):
                lines.append(f"{elem_indent}  🖼 Рисунок {elem.number or '?'}")
            elif isinstance(elem, FormulaElement):
                lines.append(f"{elem_indent}  ∑ Формула ({elem.number or '?'})")

    for child in section.children:
        lines.extend(_print_section(child, depth + 1, show_paragraphs))

    return lines


def print_headings(model: DocumentModel) -> str:
    lines = []
    lines.append("ЗАГОЛОВКИ:")
    lines.append("-" * 40)

    for h in model.headings:
        indent = "  " * (h.level - 1)
        struct = " [С]" if h.is_structural else ""
        num = f"{h.number} " if h.number else ""
        lines.append(f"[{h.index:3d}] {indent}L{h.level}: {num}{h.text[:50]}{struct}")

    return "\n".join(lines)


def print_elements(model: DocumentModel, limit: int = 50) -> str:
    lines = []
    lines.append(f"ЭЛЕМЕНТЫ (первые {limit}):")
    lines.append("-" * 40)

    type_icons = {
        ElementType.HEADING: "H",
        ElementType.PARAGRAPH: "P",
        ElementType.TABLE: "T",
        ElementType.FIGURE: "F",
        ElementType.FORMULA: "∑",
    }

    for elem in model.elements[:limit]:
        icon = type_icons.get(elem.type, "?")
        text = elem.text[:60] if elem.text else ""
        lines.append(f"[{elem.index:3d}] {icon}: {text}")

    if len(model.elements) > limit:
        lines.append(f"... и ещё {len(model.elements) - limit} элементов")

    return "\n".join(lines)
