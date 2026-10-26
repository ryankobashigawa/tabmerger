#!/bin/sh
# Regenerates sitemap.xml from every public .html page in the repo.
# lastmod is each page's last commit date (today, if it has uncommitted changes).
# Test pages are excluded. Run from anywhere: sh scripts/update-sitemap.sh
set -e
cd "$(dirname "$0")/.."
BASE="https://stackmysheets.com"
TODAY="$(date +%Y-%m-%d)"

{
  echo '<?xml version="1.0" encoding="UTF-8"?>'
  echo '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">'
  git ls-files --cached --others --exclude-standard '*.html' \
    | grep -vE '^(test|test-fixtures|vendor|scripts)/' \
    | sort \
    | while read -r file; do
        case "$file" in
          index.html) path="/" ;;
          */index.html) path="/${file%index.html}" ;;
          # GitHub Pages serves foo.html at /foo; list the extensionless URL
          # so it matches each page's canonical link.
          *) path="/${file%.html}" ;;
        esac
        if [ -n "$(git status --porcelain -- "$file")" ]; then
          lastmod="$TODAY"
        else
          lastmod="$(git log -1 --format=%cs -- "$file")"
        fi
        echo '  <url>'
        echo "    <loc>$BASE$path</loc>"
        echo "    <lastmod>${lastmod:-$TODAY}</lastmod>"
        echo '  </url>'
      done
  echo '</urlset>'
} > sitemap.xml

echo "Wrote sitemap.xml:"
grep '<loc>' sitemap.xml | sed 's/ *<\/*loc>//g'
