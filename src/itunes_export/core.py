"""Core logic: COM connection, track extraction, export formats."""

import csv
import json
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path


TRACK_FIELDS = [
    "Name", "Artist", "Album", "Album Artist", "Composer",
    "Genre", "Year", "Track Number", "Track Count",
    "Disc Number", "Disc Count", "Duration",
    "Bit Rate", "Sample Rate", "Rating", "Play Count",
    "Date Added", "Location", "BPM", "Comment", "Kind",
]


def connect():
    """Connect to iTunes/Apple Music via COM and return the application object."""
    try:
        import win32com.client  # type: ignore[import-untyped]
        return win32com.client.Dispatch("iTunes.Application")
    except ImportError as e:
        raise RuntimeError(
            "pywin32 is required on Windows.  Run:  pip install pywin32"
        ) from e
    except Exception as e:
        raise RuntimeError(
            f"Could not connect to iTunes/Apple Music COM: {e}\n"
            "Make sure Apple Music or iTunes is installed."
        ) from e


def get_user_playlists(app):
    """Return list of (name, playlist_com_object) for all user playlists."""
    results = []
    for source in app.Sources:
        if source.Kind == 1:  # kITSourceLibrary
            for pl in source.Playlists:
                results.append((pl.Name, pl))
    return results


def _safe(obj, attr):
    try:
        v = getattr(obj, attr)
        return v if v is not None else ""
    except Exception:
        return ""


def _format_duration(seconds):
    if not seconds:
        return ""
    s = int(seconds)
    return f"{s // 60}:{s % 60:02d}"


def track_row(track):
    """Extract a dict of metadata fields from a COM track object."""
    try:
        raw_date = track.DateAdded
        date_added = raw_date.strftime("%Y-%m-%d %H:%M:%S") if raw_date else ""
    except Exception:
        date_added = ""

    return {
        "Name":         _safe(track, "Name"),
        "Artist":       _safe(track, "Artist"),
        "Album":        _safe(track, "Album"),
        "Album Artist": _safe(track, "AlbumArtist"),
        "Composer":     _safe(track, "Composer"),
        "Genre":        _safe(track, "Genre"),
        "Year":         _safe(track, "Year"),
        "Track Number": _safe(track, "TrackNumber"),
        "Track Count":  _safe(track, "TrackCount"),
        "Disc Number":  _safe(track, "DiscNumber"),
        "Disc Count":   _safe(track, "DiscCount"),
        "Duration":     _format_duration(_safe(track, "Duration")),
        "Bit Rate":     _safe(track, "BitRate"),
        "Sample Rate":  _safe(track, "SampleRate"),
        "Rating":       _safe(track, "Rating"),
        "Play Count":   _safe(track, "PlayedCount"),
        "Date Added":   date_added,
        "Location":     _safe(track, "Location"),
        "BPM":          _safe(track, "BPM"),
        "Comment":      _safe(track, "Comment"),
        "Kind":         _safe(track, "KindAsString"),
    }


def list_playlists(playlists):
    """Print playlist names and track counts to stdout."""
    print(f"\nAvailable playlists ({len(playlists)}):")
    for i, (name, pl) in enumerate(playlists, 1):
        try:
            count = pl.Tracks.Count
        except Exception:
            count = "?"
        if isinstance(count, int):
            print(f"  {i:3}. {name}  ({count:,} tracks)")
        else:
            print(f"  {i:3}. {name}  ({count} tracks)")


def find_playlist(playlists, name):
    """Return matching (name, playlist) pairs; exact match first, then partial."""
    exact = [(n, pl) for n, pl in playlists if n.lower() == name.lower()]
    if exact:
        return exact
    return [(n, pl) for n, pl in playlists if name.lower() in n.lower()]


def export_csv(rows, path):
    """Write track rows to a UTF-8 CSV file."""
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=TRACK_FIELDS)
        writer.writeheader()
        writer.writerows(rows)


def export_xml(rows, playlist_name, path):
    """Write track rows to a UTF-8 XML file."""
    root = ET.Element("Playlist")
    root.set("name", playlist_name)
    root.set("exported", datetime.now().isoformat())
    root.set("count", str(len(rows)))
    for row in rows:
        el = ET.SubElement(root, "Track")
        for field, value in row.items():
            child = ET.SubElement(el, field.replace(" ", ""))
            child.text = str(value)
    tree = ET.ElementTree(root)
    ET.indent(tree, space="  ")
    tree.write(path, encoding="utf-8", xml_declaration=True)


def export_json(rows, playlist_name, path):
    """Write track rows to a UTF-8 JSON file."""
    payload = {
        "playlist": playlist_name,
        "exported": datetime.now().isoformat(),
        "count": len(rows),
        "tracks": rows,
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)
