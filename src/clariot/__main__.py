"""Command line entry point.

    python -m clariot --self-check      verify Outlook, folders and configuration
    python -m clariot --setup-folders   create the Outlook folders
    python -m clariot --preview PDF     see what a report would say
    python -m clariot                   ingest new alerts (run often)
    python -m clariot --report          weekly reports per client (run on Friday)
    python -m clariot --pending         list what is waiting for the weekly report
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from .audit import AuditLog
from .config import (
    ClientDirectory,
    ConfigError,
    load_labels,
    load_settings,
    load_value_noise,
)
from .report_builder import build_template_env
from .glossary import Glossary
from .ingest import Ingestor
from .ledger import Ledger
from .store import EventStore
from .weekly import WeeklyReporter

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def configure_logging(log_file: Path, verbose: bool) -> None:
    """Log to the file and to the console: the technician sees one, we read the other."""
    log_file.parent.mkdir(parents=True, exist_ok=True)
    handlers: list[logging.Handler] = [
        logging.FileHandler(log_file, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ]
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s %(levelname)-7s %(name)s: %(message)s",
        handlers=handlers,
    )


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """The command line. Every flag maps to one entry point below."""
    parser = argparse.ArgumentParser(prog="clariot", description=__doc__)
    parser.add_argument("--dry-run", action="store_true", help="write nothing: no drafts, no state")
    parser.add_argument(
        "--report",
        action="store_true",
        help="build the weekly report per client from what has accumulated",
    )
    parser.add_argument(
        "--pending",
        action="store_true",
        help="list the alarms waiting for the weekly report, then exit",
    )
    parser.add_argument("--limit", type=int, default=None, help="process at most N messages")
    parser.add_argument("--self-check", action="store_true", help="verify the environment and exit")
    parser.add_argument(
        "--setup-folders",
        action="store_true",
        help="create the Outlook folders named in settings.yaml, then exit",
    )
    parser.add_argument(
        "--simulate",
        metavar="PDF",
        help="drop a test alert with this PDF into the source folder, then exit",
    )
    parser.add_argument(
        "--simulate-id",
        metavar="ID",
        default="",
        help="Message-ID for --simulate; reuse the same one to test the duplicate guard",
    )
    parser.add_argument(
        "--preview",
        metavar="PDF",
        help="parse a PDF and write the resulting draft to preview.html, without Outlook or DeepL",
    )
    parser.add_argument(
        "--review-phrases",
        "--revisar-frases",
        dest="review_phrases",
        action="store_true",
        help="list every phrase the AI translated, to audit and override it",
    )
    parser.add_argument("--verbose", action="store_true", help="debug logging")
    return parser.parse_args(argv)


def self_check(settings, directory: ClientDirectory, config_dir: Path) -> int:
    """Fail loudly and specifically, so deployment problems are obvious."""
    from .adapters import build_translator
    from .adapters.outlook import OutlookClient, OutlookError

    problems: list[str] = []

    # Clients listed in clients.yaml get the To field pre-filled; any other
    # company still gets a draft, with an empty To for the technician to fill.
    print(f"Destinatarios precargados: {len(directory)} cliente(s) en clients.yaml")

    try:
        outlook = OutlookClient()
        print("Outlook               : conectado")
        for folder in settings.outlook.source_folders:
            if outlook.folder_exists(folder):
                pendientes = outlook.messages(
                    folder, settings.outlook.include_subfolders
                )
                print(f"Carpeta '{folder}': OK ({len(pendientes)} correo(s) por revisar)")
            else:
                problems.append(f"Outlook folder '{folder}' not found under the Inbox")
    except OutlookError as exc:
        problems.append(str(exc))

    provider = settings.translation.provider
    if not settings.translation.enabled:
        print("PDF adjunto           : sin traducir (translation.provider: none)")
    else:
        try:
            translator = build_translator(settings)
            print(f"PDF ({provider})       : {translator.usage()}")
        except Exception as exc:  # noqa: BLE001 - auth, config or network failure
            problems.append(f"Traduccion de PDF con '{provider}': {exc}")

    # Phrases are translated independently of the PDF: an AI Studio key needs no
    # billing, so this half can work while the other is blocked.
    phrase_provider = settings.glossary.phrase_provider
    if phrase_provider in ("", "none"):
        print("Frases del informe    : solo el glosario aprobado (sin IA)")
    else:
        try:
            from .adapters import build_text_translator
            from .glossary import Glossary

            text = build_text_translator(settings, Glossary.load(config_dir).terms)
            print(f"Frases ({phrase_provider}){' ' * max(0, 14 - len(phrase_provider))}: {text.usage()}")
        except Exception as exc:  # noqa: BLE001
            problems.append(f"Traduccion de frases con '{phrase_provider}': {exc}")

    if problems:
        print("\nProblemas encontrados:")
        for problem in problems:
            print(f"  - {problem}")
        return 1

    print("\nTodo listo.")
    return 0


def setup_folders(settings) -> int:
    """Create the folders named in settings.yaml. The only command that writes
    to the mailbox without processing anything."""
    from .adapters.outlook import OutlookClient, OutlookError

    try:
        outlook = OutlookClient()
    except OutlookError as exc:
        print(f"{exc}", file=sys.stderr)
        return 2

    planned = [(folder, "inbox") for folder in settings.outlook.source_folders]
    if settings.outlook.processed_folder:
        planned.append((settings.outlook.processed_folder, "inbox"))
    if settings.outlook.draft_folder:
        planned.append((settings.outlook.draft_folder, settings.outlook.draft_folder_parent))
    if settings.outlook.urgent_draft_folder:
        planned.append(
            (settings.outlook.urgent_draft_folder, settings.outlook.draft_folder_parent)
        )

    for path, parent in planned:
        root = "Borradores" if parent == "drafts" else "Bandeja de entrada"
        try:
            outlook.ensure_folder(path, parent)
            print(f"OK  {root} / {path}")
        except Exception as exc:  # noqa: BLE001
            print(f"ERROR  {root} / {path}: {exc}", file=sys.stderr)
            return 1

    print("\nCarpetas listas. Falta crear la regla de Outlook que mueve")
    print("los correos de Clariot a la carpeta de origen (ver README.md, seccion 5.2).")
    return 0


def simulate(settings, pdf_path: Path, message_id: str = "") -> int:
    """Inject a test alert into the source folder. Sends no email."""
    from .adapters.outlook import OutlookClient, OutlookError

    if not pdf_path.exists():
        print(f"No existe el archivo: {pdf_path}", file=sys.stderr)
        return 2

    folder = settings.outlook.source_folders[0]
    try:
        outlook = OutlookClient()
        outlook.inject_test_message(
            folder, "PRUEBA CLARIOT", [pdf_path], message_id=message_id
        )
    except OutlookError as exc:
        print(f"{exc}", file=sys.stderr)
        return 2

    print(f"Alerta de prueba creada en '{folder}', sin leer, con {pdf_path.name} adjunto.")
    if message_id:
        print(f"Message-ID: {message_id}")
    print("\nAhora:  python -m clariot --limit 1")
    print("Para deshacer: borra ese mensaje desde Outlook.")
    return 0


def review_phrases(store, glossary) -> int:
    """Show what the AI translated, so a human can audit and override it.

    Nothing the AI wrote is beyond review: anything that reads wrong goes into
    config/glossary.yaml, which wins over the cache from that moment on.
    """
    cached = store.cached_phrases()
    if not cached:
        print("Todavia no hay frases traducidas por IA.")
        print("Aparecen aca en cuanto llegue una alerta con una frase nueva.")
        return 0

    overridden = [(src, dst) for src, dst, _ in cached if glossary.lookup(src)]
    pending = [(src, dst, prov) for src, dst, prov in cached if not glossary.lookup(src)]

    print(f"\n{len(cached)} frase(s) en la cache | {len(overridden)} ya anuladas "
          f"en glossary.yaml\n")

    if pending:
        print("=" * 72)
        print("  TRADUCIDAS POR IA — revisar")
        print("=" * 72)
        for source, target, provider in pending:
            print(f"\n  EN [{provider}]  {source}")
            print(f"  ES            {target}")

    if overridden:
        print(f"\n{'=' * 72}")
        print("  YA ANULADAS POR GLOSSARY.YAML — se usa la version de Emeltec")
        print("=" * 72)
        for source, _ in overridden:
            print(f"\n  EN            {source}")
            print(f"  ES            {glossary.lookup(source)}")

    print("\n" + "-" * 72)
    print("Para corregir una traduccion, agregala a config\\glossary.yaml bajo")
    print("'phrases:' con la redaccion que prefieran. Esa version manda desde")
    print("ese momento, en los informes nuevos.")
    print("-" * 72)
    return 0


def show_pending(store, glossary, settings) -> int:
    """List what is waiting for Friday, so nobody has to guess."""
    from .aggregate import group_by_company, group_by_equipment, period_of, spanish_period

    held = store.blocked()
    if held:
        print(f"\n{'='*68}")
        print(f"  {len(held)} ALARMA(S) RETENIDA(S) — no se creo el borrador")
        print(f"{'='*68}")
        print("\nEl informe no sale hasta que estas frases esten en el glosario,")
        print("para que el cliente nunca reciba texto en ingles.\n")

        # Grouped by phrase: the same missing wording usually blocks several
        # alarms, and it only has to be translated once.
        gaps: dict[str, list[str]] = {}
        for event in held:
            for chunk in (event.blocked_reason or "").split("; "):
                if "=" in chunk:
                    gaps.setdefault(chunk.split("=", 1)[1], []).append(
                        event.machine_label
                    )

        for phrase, machines in gaps.items():
            print(f'  {phrase}')
            print(f"      bloquea a: {', '.join(sorted(set(machines)))}\n")

        print("Agregalas a config\\glossary.yaml y volve a correr el programa:")
        print("las alarmas retenidas se procesan solas.\n")

    pending = store.pending()
    if not pending:
        if not held:
            print("No hay alarmas pendientes de informar.")
        return 0

    start, end = period_of(pending)
    print(f"\n{len(pending)} alarma(s) pendientes | periodo: {spanish_period(start, end)}\n")

    for company, events in group_by_company(pending).items():
        summaries = group_by_equipment(events, glossary, settings.critical_urgencies)
        print(f"{company} — {len(events)} alarma(s) en {len(summaries)} equipo(s)")
        for summary in summaries:
            mark = "  !" if summary.repeated else "   "
            print(f"{mark} {summary.count}x  {summary.label}")
        print()

    print("Para generar los informes:  python -m clariot --report")
    return 0


def preview(pdf_path: Path, settings, labels, value_noise, directory, resolver) -> int:
    """Render the draft one PDF would produce. No Outlook, no mailbox, no state.

    The calibration loop: run it against a real report and check every field
    landed where it should before touching the mailbox.

    Uses the same builder the real flow uses, so what you see here is what the
    client would receive. A preview that rendered a different template would be
    worse than no preview at all.
    """
    from .aggregate import summarize_equipment
    from .classifier import LEVEL_NORMAL
    from .pdf_parser import parse_pdf
    from .report_builder import build_equipment_draft
    from .store import build_event

    if not pdf_path.exists():
        print(f"No existe el archivo: {pdf_path}", file=sys.stderr)
        return 2

    alert = parse_pdf(pdf_path, labels, value_noise)
    print(f"\nCampos extraidos de {pdf_path.name}:")
    for name, value in alert.fields.items():
        shown = value if len(value) <= 90 else value[:87] + "..."
        print(f"  {name:20s} : {shown}")
    if alert.is_empty:
        print("  (ninguno: revisa config/pdf_labels.yaml)")

    # Built, never stored: the preview must leave no trace.
    event = build_event(alert, message_key="preview", pdf_path=str(pdf_path))

    summary = summarize_equipment([event], resolver, settings.critical_urgencies)
    missing = resolver.apply(alert).missing
    if missing:
        print(f"\nFrases que el glosario no cubre ({len(missing)}):")
        for name, text in missing:
            print(f'  {name}:\n    "{text}"')
        print("\n  Con la IA activada se traducen solas en la corrida real.")
    else:
        print("\nTraduccion: todas las frases resueltas.")

    draft = build_equipment_draft(
        summary,
        LEVEL_NORMAL,
        directory,
        settings,
        build_template_env(PROJECT_ROOT / "templates"),
        missing_phrases=missing,
    )
    output = PROJECT_ROOT / "preview.html"
    output.write_text(draft.html_body, encoding="utf-8")

    print(f"\nAsunto      : {draft.subject}")
    print(f"Para        : {'; '.join(draft.to) or '(vacio: lo completa el tecnico)'}")
    print(f"CC          : {'; '.join(draft.cc) or '(ninguno)'}")
    print(f"Vista previa: {output}")
    return 0


def main(argv: list[str] | None = None) -> int:
    """Load the configuration, then dispatch to the command that was asked for.

    Ordered cheapest first: the commands that need no mailbox and no state run
    before anything is opened.
    """
    args = parse_args(argv)
    config_dir = PROJECT_ROOT / "config"

    try:
        settings = load_settings(PROJECT_ROOT, config_dir)
        labels = load_labels(config_dir)
        value_noise = load_value_noise(config_dir)
        directory = ClientDirectory.load(config_dir)
        glossary = Glossary.load(config_dir)
    except ConfigError as exc:
        print(f"Error de configuración: {exc}", file=sys.stderr)
        return 2

    configure_logging(settings.paths.log_file, args.verbose)

    # These two need no mailbox and no state.
    if args.setup_folders:
        return setup_folders(settings)

    if args.simulate:
        return simulate(settings, Path(args.simulate), args.simulate_id)

    if args.self_check:
        return self_check(settings, directory, config_dir)

    store = EventStore(settings.paths.events_db)

    # Three layers: the YAML overrides, the cache reuses, the API resolves what is
    # new. Without a translator the first two still work.
    from .adapters import build_text_translator
    from .resolver import PhraseResolver

    try:
        text_translator = build_text_translator(settings, glossary.terms)
    except Exception as exc:  # noqa: BLE001
        logging.getLogger("clariot").error(
            "No se pudo iniciar el traductor de frases: %s", exc
        )
        return 2
    resolver = PhraseResolver(glossary, store, text_translator)

    if args.preview:
        return preview(
            Path(args.preview), settings, labels, value_noise, directory, resolver
        )

    if args.review_phrases:
        return review_phrases(store, glossary)

    if args.pending:
        return show_pending(store, resolver, settings)

    from .adapters import build_translator
    from .adapters.outlook import OutlookClient

    try:
        outlook = OutlookClient(
            use_signature=settings.email.use_outlook_signature,
            signature_account=settings.email.sender_account,
        )
        doc_translator = build_translator(settings)
    except Exception as exc:  # noqa: BLE001
        logging.getLogger("clariot").error("No se pudo iniciar: %s", exc)
        return 2

    env = build_template_env(PROJECT_ROOT / "templates")
    audit = AuditLog(settings.paths.audit_csv)
    log = logging.getLogger("clariot")

    if args.report:
        reporter = WeeklyReporter(
            settings=settings,
            glossary=resolver,
            directory=directory,
            outlook=outlook,
            store=store,
            audit=audit,
            env=env,
            dry_run=args.dry_run,
        )
        weekly = reporter.run()
        log.info("Informe semanal: %s", weekly.summary())
        return 1 if weekly.failed else 0

    ingestor = Ingestor(
        settings=settings,
        labels=labels,
        value_noise=value_noise,
        glossary=resolver,
        directory=directory,
        outlook=outlook,
        store=store,
        translator=doc_translator,
        ledger=Ledger(settings.paths.ledger_db),
        audit=audit,
        env=env,
        dry_run=args.dry_run,
    )
    result = ingestor.run(limit=args.limit)
    log.info("Resumen: %s", result.summary())

    purged = store.purge_older_than(settings.grouping.retention_days)
    if purged:
        log.info("Se purgaron %s eventos ya informados y antiguos", purged)

    return 1 if result.failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
