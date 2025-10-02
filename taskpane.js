/* global Office Word */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log('Estate Clause Helper loaded');
        loadClauses();
        setupSearch();
    }
});

// Sample clause database (in production, fetch from API)
const clauseDatabase = [
    {
        id: 1,
        title: "Standard Executor Powers",
        category: "Executor",
        preview: "The Executor shall have full power to manage, sell, lease, or otherwise deal with any property...",
        fullText: `The Executor shall have full power and authority to manage, sell, lease, mortgage, or otherwise deal with any property forming part of my estate, without the need for court approval, and to invest and reinvest the estate assets as they see fit, with the same freedom as if they were the absolute owner thereof.`
    },
    {
        id: 2,
        title: "Trust Distribution Clause",
        category: "Trust",
        preview: "The Trustee shall distribute income and principal to the beneficiaries as follows...",
        fullText: `The Trustee shall have discretion to distribute so much of the net income and principal of the trust as the Trustee deems necessary for the health, education, maintenance, and support of the beneficiaries, taking into account their other resources known to the Trustee.`
    },
    {
        id: 3,
        title: "Residuary Estate Clause",
        category: "Distribution",
        preview: "I give all the rest, residue, and remainder of my estate...",
        fullText: `I give, devise, and bequeath all the rest, residue, and remainder of my estate, both real and personal, of whatever nature and wherever situated, which I may own or have the right to dispose of at the time of my death, to my Trustee, to be held, administered, and distributed in accordance with the terms of the trust established herein.`
    },
    {
        id: 4,
        title: "No-Contest Clause",
        category: "Protection",
        preview: "Any beneficiary who contests this will shall forfeit their entire interest...",
        fullText: `If any beneficiary under this Will shall in any manner, directly or indirectly, contest or attack this Will or any of its provisions, or seek to impair or invalidate any of its provisions, or conspire with or assist anyone attempting to do any of those things, then I specifically disinherit that person and any legacy, devise, or benefit provided for them under this Will shall be revoked and shall pass as though that person had predeceased me.`
    },
    {
        id: 5,
        title: "Guardian Appointment",
        category: "Guardian",
        preview: "I appoint [Name] as guardian of my minor children...",
        fullText: `If at my death any of my children are minors, I appoint [GUARDIAN NAME] as guardian of the person and estate of such minor children. If [GUARDIAN NAME] is unable or unwilling to serve, I appoint [ALTERNATE GUARDIAN NAME] as alternate guardian. I request that no bond or other security be required of any guardian appointed herein.`
    },
    {
        id: 6,
        title: "Simultaneous Death Provision",
        category: "Survivorship",
        preview: "If my spouse and I die simultaneously or within 30 days of each other...",
        fullText: `If my spouse and I should die simultaneously, or under circumstances that make it difficult or impossible to determine who predeceased the other, or if my spouse should die within thirty (30) days after my death, then for the purposes of this Will, my spouse shall be deemed to have predeceased me.`
    },
    {
        id: 7,
        title: "Digital Assets Clause",
        category: "Modern Assets",
        preview: "My Executor shall have authority to access, manage, and distribute my digital assets...",
        fullText: `My Executor shall have the authority to access, handle, distribute, and dispose of my digital assets, including but not limited to emails, social media accounts, digital files, cryptocurrencies, and online accounts. I authorize my Executor to access any computer, mobile device, or online account and to obtain and use passwords or other authentication credentials necessary to carry out these powers.`
    },
    {
        id: 8,
        title: "Spendthrift Trust Provision",
        category: "Trust Protection",
        preview: "No beneficiary shall have the power to anticipate, assign, or encumber their interest...",
        fullText: `No beneficiary shall have any right, power, or authority to anticipate, assign, pledge, encumber, or otherwise alienate their interest in the trust or any distribution therefrom, either in whole or in part. No such interest shall be subject to the claims of creditors or liable to attachment, execution, or other legal process. This spendthrift provision shall not prevent the Trustee from making distributions directly to providers of goods or services for the benefit of a beneficiary.`
    }
];

let allClauses = [...clauseDatabase];

function loadClauses() {
    displayClauses(allClauses);
}

function setupSearch() {
    const searchBox = document.getElementById('searchBox');
    searchBox.addEventListener('input', (e) => {
        const searchTerm = e.target.value.toLowerCase();
        const filtered = allClauses.filter(clause => 
            clause.title.toLowerCase().includes(searchTerm) ||
            clause.preview.toLowerCase().includes(searchTerm) ||
            clause.category.toLowerCase().includes(searchTerm)
        );
        displayClauses(filtered);
    });
}

function displayClauses(clauses) {
    const listContainer = document.getElementById('clauseList');
    
    if (clauses.length === 0) {
        listContainer.innerHTML = '<div class="no-results">No clauses found. Try a different search term.</div>';
        return;
    }
    
    listContainer.innerHTML = clauses.map(clause => `
        <div class="clause-item">
            <div class="clause-title">${clause.title}</div>
            <div class="clause-preview">${clause.preview}</div>
            <span class="clause-category">${clause.category}</span>
            <button class="insert-btn" onclick="insertClause(${clause.id})">
                Insert into Document
            </button>
        </div>
    `).join('');
}

async function insertClause(clauseId) {
    const clause = allClauses.find(c => c.id === clauseId);
    if (!clause) return;

    try {
        await Word.run(async (context) => {
            // Insert at current cursor position
            const selection = context.document.getSelection();
            
            // Insert the clause title as a heading
            const titleParagraph = selection.insertParagraph(clause.title, Word.InsertLocation.end);
            titleParagraph.styleBuiltIn = Word.Style.heading2;
            
            // Insert the full text
            const textParagraph = selection.insertParagraph(clause.fullText, Word.InsertLocation.end);
            textParagraph.styleBuiltIn = Word.Style.normal;
            
            // Add some spacing
            selection.insertParagraph("", Word.InsertLocation.end);
            
            await context.sync();
            
            showStatus(`✓ "${clause.title}" inserted successfully!`, 'success');
        });
    } catch (error) {
        console.error('Error inserting clause:', error);
        showStatus(`✗ Error inserting clause: ${error.message}`, 'error');
    }
}

function showStatus(message, type) {
    const statusDiv = document.getElementById('statusMessage');
    statusDiv.textContent = message;
    statusDiv.className = `status-message status-${type}`;
    statusDiv.style.display = 'block';
    
    setTimeout(() => {
        statusDiv.style.display = 'none';
    }, 3000);
}

// Make insertClause available globally
window.insertClause = insertClause;
