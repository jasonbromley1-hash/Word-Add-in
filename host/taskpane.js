// taskpane.js – interacts with Word + clause suggestion API
/* global Office, Word */
(async () => {
  await Office.onReady();
  const inputEl = document.getElementById("inputText");
  const btnSuggest = document.getElementById("btnSuggest");
  const statusEl = document.getElementById("status");
  const suggEl = document.getElementById("suggestions");
  const setStatus = (msg) => (statusEl.textContent = msg || "");
  const clearSuggestions = () => (suggEl.innerHTML = "");
  const renderSuggestions = (suggs = []) => {
    clearSuggestions();
    if (!suggs.length) { suggEl.textContent = "No suggestions."; return; }
    suggEl.innerHTML = "";
    suggs.forEach((s, idx) => {
      const div = document.createElement("div");
      div.className = "suggest-item";
      div.innerHTML = `<strong>${idx + 1}. ${s.label || "Suggestion"}</strong><br/>${s.text}`;
      div.addEventListener("click", () => insertClause(s.text));
      suggEl.appendChild(div);
    });
  };
  const insertClause = async (text) => {
    try {
      await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.insertText(text, Word.InsertLocation.replace);
        await ctx.sync();
      });
      setStatus("Inserted suggestion ✔");
    } catch (err) { console.error(err); setStatus("Failed to insert: " + err.message); }
  };
  const handleSuggest = async () => {
    setStatus("Gathering context…");
    clearSuggestions();
    let queryText = inputEl.value.trim();
    if (!queryText) {
      try {
        queryText = await Word.run(async (ctx) => { const sel = ctx.document.getSelection(); sel.load("text"); await ctx.sync(); return sel.text.trim(); });
      } catch (e) { console.warn("Unable to read selection:", e); }
    }
    if (!queryText) { setStatus("Provide text (or select in document) first."); return; }
    setStatus("Calling AI suggester…");
    try {
      const resp = await fetch("/suggest", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ query: queryText, mode: "hybrid", k: 6, show_only: false }) });
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const data = await resp.json();
      const suggs = data?.suggestions || [];
      renderSuggestions(suggs);
      setStatus(`Received ${suggs.length} suggestions.`);
    } catch (err) { console.error(err); setStatus("Error: " + err.message); }
  };
  btnSuggest.addEventListener("click", handleSuggest);
})();
