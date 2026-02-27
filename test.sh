#!/bin/bash
# The Sidebar API Test Suite
# Requires: server running on localhost:3001, task pane connected

BASE="https://localhost:3001"
CURL="curl -sk --max-time 15"
PASS=0
FAIL=0

test_endpoint() {
  local desc="$1" method="$2" path="$3" body="$4" expect="$5"
  local url="$BASE$path"
  local result
  if [ "$method" = "GET" ]; then
    result=$($CURL "$url" 2>&1)
  else
    result=$($CURL -X "$method" -H 'Content-Type: application/json' -d "$body" "$url" 2>&1)
  fi
  if echo "$result" | python3 -c "import sys,json; d=json.load(sys.stdin); exit(0 if d.get('ok') else 1)" 2>/dev/null; then
    echo "  ✅ $desc ($(echo "$result" | python3 -c "import sys,json; print(json.load(sys.stdin).get('_ms','?'))" 2>/dev/null)ms)"
    PASS=$((PASS+1))
  else
    echo "  ❌ $desc"
    echo "     $(echo "$result" | head -c 200)"
    FAIL=$((FAIL+1))
  fi
}

echo "⚖️ The Sidebar API Test Suite"
echo "═══════════════════════════"
echo ""

echo "── Health & Meta ──"
test_endpoint "Status" GET "/api/status"
test_endpoint "Help" GET "/api/help"
test_endpoint "Ping" GET "/api/ping"

echo ""
echo "── Document Reading ──"
test_endpoint "Full text" GET "/api/document"
test_endpoint "Paragraphs (range)" GET "/api/document/paragraphs?from=0&to=5"
test_endpoint "Paragraphs (compact)" GET "/api/document/paragraphs?from=0&to=5&compact=true"
test_endpoint "Stats" GET "/api/document/stats"
test_endpoint "Structure" GET "/api/document/structure"
test_endpoint "TOC" GET "/api/document/toc"
test_endpoint "HTML" GET "/api/document/html"

echo ""
echo "── Paragraphs ──"
test_endpoint "Single paragraph" GET "/api/paragraph/5"
test_endpoint "Paragraph (compact)" GET "/api/paragraph/5?compact=true"
test_endpoint "Paragraph context" GET "/api/paragraph/5/context"
test_endpoint "Paragraph context (radius=1)" GET "/api/paragraph/5/context?radius=1"

echo ""
echo "── Selection ──"
test_endpoint "Read selection" GET "/api/selection"

echo ""
echo "── Index ──"
test_endpoint "Build index" POST "/api/index/build"
test_endpoint "Get index" GET "/api/index"
test_endpoint "Headings" GET "/api/index/headings"
test_endpoint "Index range" GET "/api/index/range?from=0&to=10"

echo ""
echo "── Find ──"
test_endpoint "Find text" POST "/api/find" '{"text":"Plaintiff"}'

echo ""
echo "── Styles ──"
test_endpoint "List styles" GET "/api/styles"

echo ""
echo "── Footnotes ──"
test_endpoint "List footnotes" GET "/api/footnotes"
test_endpoint "Search footnotes" POST "/api/footnote/search" '{"text":"Mansour"}'

echo ""
echo "── Comments ──"
test_endpoint "List comments" GET "/api/comments"

echo ""
echo "── Navigation ──"
test_endpoint "Navigate to paragraph" POST "/api/navigate" '{"index":5}'
test_endpoint "Select paragraph" POST "/api/select" '{"index":5}'

echo ""
echo "── Undo ──"
test_endpoint "Undo history" GET "/api/undo/history"

echo ""
echo "── Advanced Additions ──"
test_endpoint "Document properties" GET "/api/document/properties"
test_endpoint "Section by heading index" GET "/api/section?headingIndex=0"
test_endpoint "Bulk paragraphs" POST "/api/paragraphs/bulk" '{"indices":[0,1,2,50]}'
test_endpoint "Diff paragraph" POST "/api/paragraph/diff" '{"index":5,"compareText":"Plaintiff alleges monopolization"}'

echo ""
echo "── Prompts ──"
test_endpoint "Get prompts" GET "/api/prompts"

echo ""
echo "═══════════════════════════"
echo "Results: $PASS passed, $FAIL failed"
