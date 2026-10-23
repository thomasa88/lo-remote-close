# This file is part of lo-remote-close, a tool for remotely closing
# Libreoffice documents.
#
# Copyright (c) 2025 Thomas Axelsson
# 
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

import argparse
import enum
import logging

from ooodev.loader.lo import Lo
import ooodev.loader.inst.options as ooo_options

# Type imports found by e.g. `find .venv -iname '*x_spread*'`
# from com.sun.star.sheet import XSpreadsheetDocument, XSpreadsheet
from ooodev.loader.comp.the_desktop import TheDesktop
from com.sun.star.frame import XTitle
from com.sun.star.util import XModifiable2

class ModifiedAction(enum.Enum):
    CLOSE = enum.auto(),
    SAVE = enum.auto(),
    SKIP = enum.auto(),

def main() -> None:
    argparser = argparse.ArgumentParser()
    subparsers = argparser.add_subparsers(dest='command', required=True)
    subparsers.add_parser('list', help='List open documents')
    close_parser = subparsers.add_parser('close', help='Close documents matching the given substrings')
    close_parser.add_argument('-f', '--force', action='store_true', help="Close without saving")
    close_parser.add_argument('-s', '--save', action='store_true', help="Save before closing")
    close_parser.add_argument('substrings', nargs='+', help='Substrings for the document names')
    
    args = argparser.parse_args()
    
    Lo.load_office(Lo.ConnectSocket(), opt=ooo_options.Options(log_level=logging.WARNING))
    
    match args.command:
        case 'list':
            list_docs()
        case 'close':
            mod_action = ModifiedAction.SKIP
            if args.save:
                mod_action = ModifiedAction.SAVE
            elif args.force:
                mod_action = ModifiedAction.CLOSE
            close_docs(args.substrings, mod_action)

def list_docs() -> None:
    the_desktop = TheDesktop.from_lo()
    # I have seen the same title appear multiple times, so we use a
    # set to remove the duplicates.
    titles = set()
    for comp in the_desktop.get_components():
        # print(c) gives supported services and interfaces. The interfaces can be casted to
        # using Lo.qi().
        title_c = Lo.qi(XTitle, comp)
        if title_c:
            title = title_c.getTitle()
            mod_c = Lo.qi(XModifiable2, comp)
            if mod_c.isModified():
                title += " (Modified)"
            titles.add(title)

    titles = sorted(titles, key=str.casefold)
    print("\n".join(titles))

def close_docs(substrings: list[str], mod_action: ModifiedAction) -> None:
    assert isinstance(substrings, list)
    substr_lower = [s.lower() for s in substrings]
    the_desktop = TheDesktop.from_lo()
    for comp in the_desktop.get_components():
        title_c = Lo.qi(XTitle, comp)
        if not title_c:
            continue
        title = title_c.getTitle()
        title_lower = title.lower()
        if any([True for s in substr_lower if s in title_lower]):
            mod_c = Lo.qi(XModifiable2, comp)
            if mod_c.isModified():
                match mod_action:
                    case ModifiedAction.SAVE:
                        try:
                            Lo.save(comp)
                            print(f"Saved \"{title}\".")
                        except Exception as e:
                            print(f"Unable to save \"{title}\". Not closing: {e}")
                            continue
                    case ModifiedAction.SKIP:
                        print(f"\"{title}\" is modified. Not closing.")
                        continue
                    case ModifiedAction.CLOSE:
                        pass
            # Just assuming that all the documents are XCloseable
            Lo.close(comp)
            print(f"Closed \"{title}\".")

if __name__ == '__main__':
    main()
