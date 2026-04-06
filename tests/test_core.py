"""Tests for itunes_export core logic — uses mocked COM objects."""

import json
import tempfile
from pathlib import Path
from unittest.mock import MagicMock, patch, PropertyMock
import pytest

from itunes_export.core import (
    _safe,
    _format_duration,
    track_row,
    find_playlist,
    list_playlists,
    export_csv,
    export_xml,
    export_json,
    TRACK_FIELDS,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def make_track(name="Song", artist="Artist", album="Album", **kwargs):
    """Build a mock COM track object."""
    t = MagicMock()
    t.Name = name
    t.Artist = artist
    t.Album = album
    t.AlbumArtist = kwargs.get("album_artist", "")
    t.Composer = kwargs.get("composer", "")
    t.Genre = kwargs.get("genre", "Rock")
    t.Year = kwargs.get("year", 2020)
    t.TrackNumber = kwargs.get("track_number", 1)
    t.TrackCount = kwargs.get("track_count", 10)
    t.DiscNumber = kwargs.get("disc_number", 1)
    t.DiscCount = kwargs.get("disc_count", 1)
    t.Duration = kwargs.get("duration", 180)
    t.BitRate = kwargs.get("bit_rate", 256)
    t.SampleRate = kwargs.get("sample_rate", 44100)
    t.Rating = kwargs.get("rating", 80)
    t.PlayedCount = kwargs.get("play_count", 5)
    t.DateAdded = kwargs.get("date_added", None)
    t.Location = kwargs.get("location", "C:\\Music\\song.mp3")
    t.BPM = kwargs.get("bpm", 120)
    t.Comment = kwargs.get("comment", "")
    t.KindAsString = kwargs.get("kind", "MPEG audio file")
    return t


def make_playlist(name, tracks):
    """Build a mock COM playlist with a list of track objects."""
    pl = MagicMock()
    pl.Name = name
    pl.Tracks.Count = len(tracks)
    pl.Tracks.Item = lambda i: tracks[i - 1]
    return pl


# ---------------------------------------------------------------------------
# _safe
# ---------------------------------------------------------------------------

class TestSafe:
    def test_returns_attribute(self):
        obj = MagicMock()
        obj.Foo = "bar"
        assert _safe(obj, "Foo") == "bar"

    def test_returns_empty_on_exception(self):
        obj = MagicMock()
        type(obj).Broken = PropertyMock(side_effect=Exception("oops"))
        assert _safe(obj, "Broken") == ""

    def test_returns_empty_for_none(self):
        obj = MagicMock()
        obj.Val = None
        assert _safe(obj, "Val") == ""


# ---------------------------------------------------------------------------
# _format_duration
# ---------------------------------------------------------------------------

class TestFormatDuration:
    def test_zero(self):
        assert _format_duration(0) == ""

    def test_empty(self):
        assert _format_duration("") == ""

    def test_seconds(self):
        assert _format_duration(90) == "1:30"

    def test_exact_minute(self):
        assert _format_duration(60) == "1:00"

    def test_long(self):
        assert _format_duration(3661) == "61:01"

    def test_pads_seconds(self):
        assert _format_duration(65) == "1:05"


# ---------------------------------------------------------------------------
# track_row
# ---------------------------------------------------------------------------

class TestTrackRow:
    def test_basic_fields(self):
        t = make_track(name="Hello", artist="World", genre="Pop")
        row = track_row(t)
        assert row["Name"] == "Hello"
        assert row["Artist"] == "World"
        assert row["Genre"] == "Pop"

    def test_all_fields_present(self):
        row = track_row(make_track())
        for field in TRACK_FIELDS:
            assert field in row

    def test_duration_formatted(self):
        t = make_track(duration=125)
        row = track_row(t)
        assert row["Duration"] == "2:05"

    def test_date_added_none(self):
        t = make_track(date_added=None)
        row = track_row(t)
        assert row["Date Added"] == ""

    def test_date_added_formatted(self):
        from datetime import datetime
        t = make_track(date_added=datetime(2023, 6, 15, 10, 30, 0))
        row = track_row(t)
        assert row["Date Added"] == "2023-06-15 10:30:00"


# ---------------------------------------------------------------------------
# find_playlist
# ---------------------------------------------------------------------------

class TestFindPlaylist:
    def setup_method(self):
        self.pl_a = MagicMock()
        self.pl_b = MagicMock()
        self.pl_c = MagicMock()
        self.playlists = [
            ("My Favorites", self.pl_a),
            ("Road Trip", self.pl_b),
            ("Jazz Classics", self.pl_c),
        ]

    def test_exact_match(self):
        result = find_playlist(self.playlists, "Road Trip")
        assert len(result) == 1
        assert result[0][0] == "Road Trip"

    def test_case_insensitive_exact(self):
        result = find_playlist(self.playlists, "road trip")
        assert len(result) == 1

    def test_partial_match(self):
        result = find_playlist(self.playlists, "Jazz")
        assert len(result) == 1
        assert result[0][0] == "Jazz Classics"

    def test_no_match(self):
        result = find_playlist(self.playlists, "Nonexistent")
        assert result == []

    def test_multiple_partial_matches(self):
        playlists = [("Rock Hits", MagicMock()), ("Rock Classics", MagicMock())]
        result = find_playlist(playlists, "Rock")
        assert len(result) == 2

    def test_exact_wins_over_partial(self):
        playlists = [("Jazz", MagicMock()), ("Jazz Classics", MagicMock())]
        result = find_playlist(playlists, "Jazz")
        assert len(result) == 1
        assert result[0][0] == "Jazz"


# ---------------------------------------------------------------------------
# list_playlists
# ---------------------------------------------------------------------------

class TestListPlaylists:
    def test_prints_without_error(self, capsys):
        pl = MagicMock()
        pl.Tracks.Count = 42
        list_playlists([("My Mix", pl)])
        out = capsys.readouterr().out
        assert "My Mix" in out
        assert "42" in out

    def test_handles_count_exception(self, capsys):
        pl = MagicMock()
        type(pl.Tracks).Count = PropertyMock(side_effect=Exception)
        list_playlists([("Broken", pl)])
        out = capsys.readouterr().out
        assert "Broken" in out


# ---------------------------------------------------------------------------
# export_csv
# ---------------------------------------------------------------------------

class TestExportCsv:
    def test_creates_file(self):
        rows = [track_row(make_track(name="A")), track_row(make_track(name="B"))]
        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as f:
            path = Path(f.name)
        export_csv(rows, path)
        assert path.exists()
        content = path.read_text(encoding="utf-8-sig")
        assert "Name" in content
        assert "Artist" in content

    def test_row_count(self):
        rows = [track_row(make_track()) for _ in range(5)]
        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as f:
            path = Path(f.name)
        export_csv(rows, path)
        lines = path.read_text(encoding="utf-8-sig").strip().splitlines()
        assert len(lines) == 6  # header + 5 rows

    def test_track_name_in_output(self):
        rows = [track_row(make_track(name="Stairway To Heaven"))]
        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as f:
            path = Path(f.name)
        export_csv(rows, path)
        assert "Stairway To Heaven" in path.read_text(encoding="utf-8-sig")


# ---------------------------------------------------------------------------
# export_xml
# ---------------------------------------------------------------------------

class TestExportXml:
    def test_creates_valid_xml(self):
        rows = [track_row(make_track(name="TestSong"))]
        with tempfile.NamedTemporaryFile(suffix=".xml", delete=False) as f:
            path = Path(f.name)
        export_xml(rows, "Test Playlist", path)
        import xml.etree.ElementTree as ET
        tree = ET.parse(path)
        root = tree.getroot()
        assert root.tag == "Playlist"
        assert root.get("name") == "Test Playlist"
        assert root.get("count") == "1"

    def test_track_name_present(self):
        rows = [track_row(make_track(name="MyTrack"))]
        with tempfile.NamedTemporaryFile(suffix=".xml", delete=False) as f:
            path = Path(f.name)
        export_xml(rows, "PL", path)
        content = path.read_text(encoding="utf-8")
        assert "MyTrack" in content

    def test_multiple_tracks(self):
        rows = [track_row(make_track()) for _ in range(3)]
        with tempfile.NamedTemporaryFile(suffix=".xml", delete=False) as f:
            path = Path(f.name)
        export_xml(rows, "Multi", path)
        import xml.etree.ElementTree as ET
        root = ET.parse(path).getroot()
        assert len(list(root)) == 3


# ---------------------------------------------------------------------------
# export_json
# ---------------------------------------------------------------------------

class TestExportJson:
    def test_creates_valid_json(self):
        rows = [track_row(make_track(name="JsonSong"))]
        with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as f:
            path = Path(f.name)
        export_json(rows, "JSON Playlist", path)
        data = json.loads(path.read_text(encoding="utf-8"))
        assert data["playlist"] == "JSON Playlist"
        assert data["count"] == 1
        assert len(data["tracks"]) == 1
        assert data["tracks"][0]["Name"] == "JsonSong"

    def test_multiple_tracks(self):
        rows = [track_row(make_track(name=f"Track {i}")) for i in range(5)]
        with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as f:
            path = Path(f.name)
        export_json(rows, "Five", path)
        data = json.loads(path.read_text(encoding="utf-8"))
        assert data["count"] == 5
        assert len(data["tracks"]) == 5

    def test_has_exported_timestamp(self):
        rows = [track_row(make_track())]
        with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as f:
            path = Path(f.name)
        export_json(rows, "PL", path)
        data = json.loads(path.read_text(encoding="utf-8"))
        assert "exported" in data
        assert len(data["exported"]) > 0
