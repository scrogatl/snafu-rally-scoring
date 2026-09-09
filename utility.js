/*
 * snafu-rally-scoring
 * Copyright (C) 2026 Scott Rogers
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

// Last edited: 2026-09-08
/**
 * Deletes every rider score sheet, Leader Board, AND Master Scoring, leaving
 * only Config/Rider Master/Bonus Master. Leader Board and Master Scoring are
 * deleted along with the rider sheets - not kept - since both are derived/
 * regenerable data exactly like they are: re-running setup() recreates all
 * three from scratch (fresh rows/columns for every current rider and bonus)
 * the same way it recreates any rider sheet that's missing. Reads sheet names
 * from Config where available (so it stays in sync with your actual sheet_*
 * settings instead of a second hardcoded copy of them), falling back to the
 * documented defaults if Config can't be loaded (e.g. it doesn't exist, or is
 * missing some keys) - this utility should still work even before Config is
 * fully set up.
 */
function deleteAllRiderSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let config = {};
  try { config = loadConfig(ss); } catch (e) { Logger.log('deleteAllRiderSheets: could not load Config (' + e.message + ') - using default sheet names.'); }

  const keepSheets = [
    'Config',
    config['sheet_rider_master'] || 'Rider Master',
    config['sheet_bonus_master'] || 'Bonus Master',
  ];

  const sheets = ss.getSheets();
  for (const sheet of sheets) {
    const name = sheet.getName();
    if (!keepSheets.includes(name)) {
      ss.deleteSheet(sheet);
      Logger.log('Deleted: ' + name);
    }
  }
  Logger.log('Done.');
}
