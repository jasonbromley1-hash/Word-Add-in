// taskpane.js – interacts with Word + clause suggestion API
/* global Office, Word */

// CONFIGURATION: Update this with your actual backend API URL
const API_BASE_URL = "https://YOUR-FUNCTION-NAME.azurewebsites.net/api"; // ⚠️ CHANGE THIS!
// Examples:
// - Azure Function: "https://estate-clause-api.azurewebsites.net/api"
// - AWS Lambda: "https://your-api-id.execute-api.region.amazonaws.com/prod"
// - Custom server: "https://api.yourfirm.com"

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
    if (!suggs.length) { 
      suggEl.textContent = "No suggestions."; 
      return; 
    }
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
      setStatus("Inserted suggestion ✓");
    } catch (err) { 
      console.error(err); 
      setStatus("Failed to insert: " + err.message); 
    }
  };
  
  const handleSuggest = async () => {
    setStatus("Gathering context…");
    clearSuggestions();
    
    let queryText = inputEl.value.trim();
    
    // If no text in input, try to get selected text from document
    if (!queryText) {
      try {
        queryText = await Word.run(async (ctx) => { 
          const sel = ctx.document.getSelection(); 
          sel.load("text"); 
          await ctx.sync(); 
          return sel.text.trim(); 
        });
      } catch (e) { 
        console.warn("Unable to read selection:", e); 
      }
    }
    
    if (!queryText) { 
      setStatus("Provide text (or select in document) first."); 
      return; 
    }
    
    setStatus("Calling AI suggester…");
    
    try {
      // Construct the full API endpoint URL
      const apiUrl = `${API_BASE_URL}/suggest`;
      
      const resp = await fetch(apiUrl, { 
        method: "POST", 
        headers: { 
          "Content-Type": "application/json" 
        }, 
        body: JSON.stringify({ 
          query: queryText, 
          mode: "hybrid", 
          k: 6, 
          show_only: false 
        }) 
      });
      
      if (!resp.ok) {
        throw new Error(`HTTP ${resp.status}: ${resp.statusText}`);
      }
      
      const data = await resp.json();
      const suggs = data?.suggestions || [];
      
      renderSuggestions(suggs);
      setStatus(`Received ${suggs.length} suggestions.`);
      
    } catch (err) { 
      console.error(err); 
      setStatus("Error: " + err.message + " - Check if backend API is running and API_BASE_URL is correct."); 
    }
  };
  
  btnSuggest.addEventListener("click", handleSuggest);
})();
window.insertClause = insertClause;

