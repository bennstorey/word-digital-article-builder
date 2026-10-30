#!/bin/sh
# publish-plugin.sh — publish the Studio plug-in to GitHub Pages.
#
# Two addresses, so that work on Discovery never moves the lab:
#   word-digital-plugin.js      the lab loads this; it changes only on `promote`
#   dev/word-digital-plugin.js  Discovery loads this; `dev` rebuilds it
#
#   publish-plugin.sh dev       build, commit and push dev/. Refuses if the push would also change the lab's file.
#   publish-plugin.sh promote   copy the dev file as published, byte for byte, over the lab's file, and push that.
#
# A push also publishes index.html, the standalone web page, so that page moves on `dev`.
set -e
cd "$(dirname "$0")"
BASE=https://bennstorey.github.io/word-digital-article-builder
FILE=word-digital-plugin.js

committed() { git show "HEAD:$1" 2>/dev/null | shasum -a 256 | cut -d' ' -f1; }
served() { curl -fsS -m 30 "$BASE/$1?t=$(date +%s)" 2>/dev/null | shasum -a 256 | cut -d' ' -f1; }
EMPTY=$(printf '' | shasum -a 256 | cut -d' ' -f1)   # what a missing file or a failed fetch hashes to: two of those are not "the same"
same() { a="$(committed "$1")"; [ "$a" != "$EMPTY" ] && [ "$a" = "$(served "$1")" ]; }

# GitHub Pages serves main, and both modes push main: from another branch the build would never arrive.
case "$1" in dev|promote) [ "$(git branch --show-current)" = main ] || { echo "Refusing: this publishes main, and you are on $(git branch --show-current). Nothing published."; exit 1; } ;; esac

case "$1" in
  dev)
    git diff --quiet HEAD -- index.html plugin-shell.js build-plugin.js || { echo "Refusing: commit the source first, so the dev build can be traced to it. Nothing published."; exit 1; }
    same "$FILE" || { echo "Refusing: the committed $FILE is not the file the lab loads, so this push would change the lab. Nothing published."; exit 1; }
    node build-plugin.js
    git add "dev/$FILE"
    git diff --cached --quiet -- "dev/$FILE" || git commit -q -m "Plug-in dev build $(date -u +%Y-%m-%dT%H:%M:%SZ)" -- "dev/$FILE"
    git push origin main
    echo "published for Discovery: $BASE/dev/$FILE (Pages takes about a minute). The lab's file is unchanged."
    ;;
  promote)
    same "dev/$FILE" || { echo "Refusing: the committed dev/$FILE is not the file Discovery loads, so it has not been proven there. Nothing published."; exit 1; }
    git show "HEAD:dev/$FILE" > "$FILE"
    git add "$FILE"
    git diff --cached --quiet -- "$FILE" || git commit -q -m "Plug-in promoted to the lab $(date -u +%Y-%m-%dT%H:%M:%SZ)" -- "$FILE"
    git push origin main
    echo "published for the lab: $BASE/$FILE (Pages takes about a minute)"
    ;;
  *)
    echo "usage: publish-plugin.sh dev | promote"; exit 2
    ;;
esac
