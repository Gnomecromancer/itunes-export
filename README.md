# itunes-export

Export Apple Music / iTunes playlists to CSV, XML, or JSON from the command line — no manual dragging, no XML library hacks. Uses the iTunes COM interface directly.

```
pip install itunes-export
```

**Requires Windows with Apple Music or iTunes installed.**

## Usage

```
itunes-export --list
itunes-export --playlist "My Favorites"
itunes-export --playlist "Road Trip" --format xml
itunes-export --playlist "Road Trip" --format json
itunes-export --playlist "Road Trip" --output road_trip.csv
```

## Exported fields

Name, Artist, Album, Album Artist, Composer, Genre, Year, Track Number, Track Count, Disc Number, Disc Count, Duration, Bit Rate, Sample Rate, Rating, Play Count, Date Added, Location, BPM, Comment, Kind

## Formats

- **CSV** (default) — UTF-8 with BOM for Excel compatibility
- **XML** — structured `<Playlist><Track>` document
- **JSON** — `{ playlist, exported, count, tracks: [...] }`

## Python API

```python
from itunes_export import connect, get_user_playlists, find_playlist, export_csv

app = connect()
playlists = get_user_playlists(app)
matches = find_playlist(playlists, "Road Trip")
name, pl = matches[0]

rows = [track_row(pl.Tracks.Item(i)) for i in range(1, pl.Tracks.Count + 1)]
export_csv(rows, "road_trip.csv")
```

## Requirements

- Windows
- Apple Music or iTunes installed
- Python 3.8+

## License

MIT
