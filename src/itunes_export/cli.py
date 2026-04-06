"""Command-line interface for itunes-export."""

import sys
import argparse
from pathlib import Path

from .core import (
    connect,
    get_user_playlists,
    find_playlist,
    list_playlists,
    track_row,
    export_csv,
    export_xml,
    export_json,
)


def _collect_tracks(pl, name):
    """Pull all track rows from a playlist with a progress bar."""
    try:
        tracks = pl.Tracks
        total = tracks.Count
    except Exception as e:
        print(f"Error accessing tracks: {e}", file=sys.stderr)
        sys.exit(1)

    print(f"\nExporting '{name}'  ({total:,} tracks) ...")
    rows = []
    for i in range(1, total + 1):
        rows.append(track_row(tracks.Item(i)))
        if i % 50 == 0 or i == total:
            pct = i / total * 100
            bar = "#" * int(pct // 2) + "-" * (50 - int(pct // 2))
            print(f"\r  [{bar}] {i:,}/{total:,} ({pct:.0f}%)", end="", flush=True)
    print()
    return rows


def main():
    parser = argparse.ArgumentParser(
        description="Export Apple Music / iTunes playlists to CSV, XML, or JSON",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  itunes-export --list
  itunes-export --playlist "My Favorites"
  itunes-export --playlist "Road Trip" --format xml
  itunes-export --playlist "Road Trip" --format json
  itunes-export --playlist "Road Trip" --output road_trip.csv
        """,
    )
    parser.add_argument("--list", "-l", action="store_true",
                        help="List all playlists with track counts")
    parser.add_argument("--playlist", "-p", metavar="NAME",
                        help="Playlist to export (partial name match ok)")
    parser.add_argument("--format", "-f", choices=["csv", "xml", "json"], default="csv",
                        help="Output format (default: csv)")
    parser.add_argument("--output", "-o", metavar="FILE",
                        help="Output path (default: <playlist name>.<format>)")

    args = parser.parse_args()
    if not args.list and not args.playlist:
        parser.print_help()
        sys.exit(0)

    print("Connecting to Apple Music / iTunes ...")
    try:
        app = connect()
    except RuntimeError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    playlists = get_user_playlists(app)

    if args.list:
        list_playlists(playlists)

    if args.playlist:
        matches = find_playlist(playlists, args.playlist)
        if not matches:
            print(f"Error: Playlist '{args.playlist}' not found.", file=sys.stderr)
            print("Use --list to see available playlists.", file=sys.stderr)
            sys.exit(1)
        if len(matches) > 1:
            print(f"Multiple playlists match '{args.playlist}':", file=sys.stderr)
            for n, _ in matches:
                print(f"  - {n}", file=sys.stderr)
            print("Please be more specific.", file=sys.stderr)
            sys.exit(1)

        name, pl = matches[0]
        rows = _collect_tracks(pl, name)

        if args.output is None:
            safe_name = "".join(c for c in name if c.isalnum() or c in " _-").strip()
            out_path = Path(f"{safe_name}.{args.format}")
        else:
            out_path = Path(args.output)

        if args.format == "csv":
            export_csv(rows, out_path)
        elif args.format == "xml":
            export_xml(rows, name, out_path)
        else:
            export_json(rows, name, out_path)

        print(f"Done. {len(rows):,} tracks → {out_path.resolve()}")
