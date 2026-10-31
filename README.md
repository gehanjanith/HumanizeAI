# HumanizeAI

HumanizeAI lightly edits UK curriculum specification Word documents so unit titles, aims, learning outcomes, and assessment criteria read more naturally. It uses the DeepSeek `deepseek-chat` API and checks rewritten content to help preserve meaning, numbers, and level-appropriate action verbs.

## Requirements

- Python 3
- A DeepSeek API key for rewriting (not needed for a dry run)

## Setup

1. Install the dependencies:

   ```powershell
   python -m pip install -r requirements.txt
   ```

2. Copy `.env.example` to `.env` and set your API key:

   ```text
   DEEPSEEK_API_KEY=your-deepseek-api-key
   ```

   Keep `.env` private; it is excluded from Git.

3. Place one or more `.docx` files in `Input/`.

## Usage

Run a preview that lists the text targets without making API calls:

```powershell
python main.py --dry-run
```

Rewrite all eligible documents in `Input/`:

```powershell
python main.py
```

The script writes results to `output/` using the input filenames. Existing results are skipped unless `--overwrite` is specified.

### Options

| Option | Description |
| --- | --- |
| `--file NAME` | Process one file from `Input/`, for example `--file "specification.docx"`. |
| `--dry-run` | List detected rewrite targets without calling the API or saving rewritten documents. |
| `--limit N` | Process only the first N targets. |
| `--overwrite` | Replace an existing output document. |
| `--level N` | Override the RQF level used for verb guidance. Supported levels: 3, 4, 5, and 7. By default, the level is detected from the filename and falls back to 5. |

Options can be combined, for example:

```powershell
python main.py --file "specification.docx" --dry-run --limit 10
python main.py --file "specification.docx" --level 5 --overwrite
```

## Results and limitations

- Only unit titles, unit aims, learning outcomes, and assessment criteria in recognized unit tables are considered.
- Rewrites rejected by the script's validation checks are left unchanged.
- A `.log.txt` file is written beside each processed output document. It records rewrite decisions and any errors.
- API use may incur charges and requires an internet connection.
- Review generated documents before relying on them; automated rewriting is not a substitute for academic or qualification-level review.
