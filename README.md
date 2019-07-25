# perfect 😎

Process Excel files with Clojure

## Disclaimer

This project, though very cool, has unfortunately slipped down in priority because of my work on [panthera](https://github.com/alanmarazzi/panthera). I plan to get back to this ASAP because I actually need this, but if in the meantime you'd like to try it out and contribute ideas (check [this issue](https://github.com/alanmarazzi/perfect/issues/1) to understand some of the angles we're trying to tackle with this), code and so on, please feel free to pop a pull request! :rocket:

## Usage

Very light API at the moment:

```clojure
(require '[perfect.reader :refer [read-workbook]])
(read-workbook "myfile.xlsx" :method :full :sheetid 0)
```

## License

Copyright © 2019 Alan Marazzi

This program and the accompanying materials are made available under the
terms of the Eclipse Public License 2.0 which is available at
http://www.eclipse.org/legal/epl-2.0.

This Source Code may also be made available under the following Secondary
Licenses when the conditions for such availability set forth in the Eclipse
Public License, v. 2.0 are satisfied: GNU General Public License as published by
the Free Software Foundation, either version 2 of the License, or (at your
option) any later version, with the GNU Classpath Exception which is available
at https://www.gnu.org/software/classpath/license.html.
