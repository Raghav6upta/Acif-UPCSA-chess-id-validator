
CHESS PLAYER VERIFIER — v0.1

What this prototype does
- Opens a local web interface.
- Lets the user choose AICF or UPCSA.
- Accepts the teacher's Excel with a PlayersID / Player ID column.
- Preserves the original worksheet and appends fetched data.
- AICF: calls https://admin.aicf.in/api/players and exact-matches aicf_id.
- UPCSA: fetches https://www.upchess.org/view_player.php?id=... from the local Python process and parses the player page.
- Caches duplicate player IDs during one run.

Important
This is a DEVELOPMENT prototype, not the final packaged teacher-facing app.
The UPCSA parser is deliberately written with fallbacks and should be tested against a representative batch of real player pages before packaging.

How to run
1. Install Python 3.11+.
2. Open a terminal in this folder.
3. Run:
   pip install -r requirements.txt
4. Run:
   python app.py
5. A browser window should open automatically.

