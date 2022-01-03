"""Parts.

### (Draft)
The protocol of `Part` .

* It has `shape` attribute which represents `Shape`.
    - Maybe instead of `shape`, `shapes` is used, for example,  `Markdown` has `shapes` since `Table` and `Textbox` cannot be grouped.
* `make` can generate it's from given `script`
* `script` attributes return `Text` format of `Shape`.
* `compile` changes of the `Content` of
"""

from fairypptx.parts.latex import Latex  # NOQA 
