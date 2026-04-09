#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$repo_root"

echo "[release-check] verifying published scope"
unexpected_files=()
while IFS= read -r file; do
  case "$file" in
    src/*|test/*|scripts/*|.github/workflows/*)
      ;;
    package.json|package-lock.json|.gitignore|.gitattributes|build_reference_library.js|report_runner.js)
      ;;
    assets/.gitkeep|reference_library/.gitkeep|reference_library/libraries/.gitkeep|reference_library/master/.gitkeep|reference_library/work_meeting/.gitkeep)
      ;;
    *)
      unexpected_files+=("$file")
      ;;
  esac
done < <(git ls-files)

if (( ${#unexpected_files[@]} > 0 )); then
  echo "Unexpected tracked files found:"
  printf '%s\n' "${unexpected_files[@]}"
  exit 1
fi

echo "[release-check] checking JavaScript syntax"
git ls-files '*.js' '*.cjs' '*.mjs' | xargs -r -n 1 node --check

echo "[release-check] running test suite"
node --test test/layoutVerification.test.js
node --test test/workflowSessionService.test.js
