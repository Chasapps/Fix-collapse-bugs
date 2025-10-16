// ============================================================================
// CATEGORY PICKER MODAL - Interactive Category Selection Component
// ============================================================================
// This module provides a searchable, paginated modal dialog for selecting categories.
// It's used when clicking the "+" button next to a transaction to assign a category.
//
// KEY CONCEPTS FOR BEGINNERS:
// - IIFE (Immediately Invoked Function Expression): The (function(){...})() pattern
//   creates a private scope so our variables don't pollute the global namespace
// - Pagination: Splitting a long list into pages to improve performance and UX
// - Event Delegation: Attaching event listeners to parent elements
// ============================================================================

(function() {
  // ============================================================================
  // CONFIGURATION AND STATE
  // ============================================================================
  
  // How many categories to show per page
  const PAGE_SIZE = 15;
  
  // Current state variables (private to this module)
  let currentPage = 1;           // Which page we're currently viewing
  let allCategories = [];        // Complete list of all available categories
  let filteredCategories = [];   // Categories after search filtering
  
  // ============================================================================
  // HTML TEMPLATE
  // ============================================================================
  
  // This is the complete HTML structure for the modal
  // We create it once and reuse it each time the modal opens
  const tpl = `
  <div class="catpicker-backdrop" id="catpickerBackdrop" role="dialog" aria-modal="true" aria-labelledby="catpickerTitle">
    <div class="catpicker-dialog">
      
      <!-- Modal Header -->
      <div class="catpicker-header">
        <h2 class="catpicker-title" id="catpickerTitle">Pick a category</h2>
      </div>
      
      <!-- Modal Body -->
      <div class="catpicker-body">
        <!-- Search input to filter categories -->
        <input id="catpickerSearch" class="catpicker-search" type="text" placeholder="Search categories…" />
        
        <!-- List container for category items -->
        <div id="catpickerList" class="catpicker-list" role="listbox" aria-label="Categories"></div>
        
        <!-- Pagination controls -->
        <div id="catpickerPager" class="catpicker-pager"></div>
      </div>
      
      <!-- Modal Actions (buttons) -->
      <div class="catpicker-actions">
        <button class="catpicker-btn" id="catpickerCancel">Cancel</button>
        <button class="catpicker-btn primary" id="catpickerUse">Use category</button>
      </div>
      
    </div>
  </div>`;

  // ============================================================================
  // MODAL INITIALIZATION
  // ============================================================================
  
  /**
   * Ensures the modal HTML exists in the page
   * Only creates it once, even if called multiple times
   */
  function ensureModal() {
    // Check if modal already exists
    if (document.getElementById('catpickerBackdrop')) return;
    
    // Create a temporary container
    const wrap = document.createElement('div');
    wrap.innerHTML = tpl;
    
    // Add the modal to the page body
    document.body.appendChild(wrap.firstElementChild);
  }

  // ============================================================================
  // CATEGORY LIST RENDERING
  // ============================================================================
  
  /**
   * Builds the list of category items in the modal
   * @param {HTMLElement} el - The container element to populate
   * @param {Array} cats - Array of category names to display
   * @param {string} picked - Currently selected category
   */
  function buildList(el, cats, picked) {
    // Clear existing content
    el.innerHTML = '';
    
    /**
     * Factory function to create a single category item element
     * @param {string} name - Category name
     * @returns {HTMLElement} The category item element
     */
    const mk = (name) => {
      // Create the item container
      const div = document.createElement('div');
      div.className = 'catpicker-item';
      
      // ARIA attributes for accessibility (screen readers)
      div.setAttribute('role', 'option');
      div.dataset.name = name;
      
      // Mark as selected if this is the currently picked category
      if (name === picked) div.setAttribute('aria-selected', 'true');
      
      // Category name
      const span = document.createElement('span');
      span.textContent = name;
      
      // Badge (currently empty, could show usage count)
      const badge = document.createElement('span');
      badge.className = 'catpicker-badge';
      badge.textContent = '';
      
      // Append children
      div.appendChild(span);
      div.appendChild(badge);
      
      // Click handler for this category item
      div.addEventListener('click', () => {
        const nm = (div.dataset.name || '').toLowerCase().trim();
        
        // Check if this is the special "Add new category" option
        const isAdd = nm.startsWith('+') || nm.startsWith('➕') || nm.indexOf('add new category') !== -1;
        
        if (isAdd) {
          // For "Add new" option, immediately trigger the "Use" button
          div.setAttribute('aria-selected', 'true');
          try {
            document.getElementById('catpickerUse').click();
          } catch(e) {}
          return;
        }
        
        // For regular categories, update selection state
        // First, deselect all items
        document.querySelectorAll('.catpicker-item[aria-selected="true"]')
          .forEach(x => x.removeAttribute('aria-selected'));
        
        // Then select this item
        div.setAttribute('aria-selected', 'true');
      });
      
      return div;
    };
    
    // Create and append an item for each category
    cats.forEach(c => el.appendChild(mk(c)));
  }

  // ============================================================================
  // PAGINATION
  // ============================================================================
  
  /**
   * Renders a specific page of categories
   * @param {number} page - Page number to render (1-based)
   * @param {Array} categories - Full list of categories to paginate
   * @param {string} picked - Currently selected category
   */
  function renderPage(page, categories, picked) {
    // Calculate total pages
    const totalPages = Math.max(1, Math.ceil(categories.length / PAGE_SIZE));
    
    // Validate page number
    if (page < 1) page = 1;
    if (page > totalPages) page = totalPages;
    
    currentPage = page;
    
    // Calculate which items to show on this page
    const start = (page - 1) * PAGE_SIZE;
    const pageItems = categories.slice(start, start + PAGE_SIZE);
    
    // Render the items
    const list = document.getElementById('catpickerList');
    buildList(list, pageItems, picked);
    
    // Render pagination controls
    renderPager(totalPages, categories.length);
  }

  /**
   * Renders the pagination controls (Prev/Next buttons and page info)
   * @param {number} totalPages - Total number of pages
   * @param {number} totalItems - Total number of items (for display)
   */
  function renderPager(totalPages, totalItems) {
    const pager = document.getElementById('catpickerPager');
    if (!pager) return;
    
    // If only one page, hide pagination
    if (totalPages <= 1) {
      pager.innerHTML = '';
      return;
    }
    
    // Calculate item range for current page
    const start = (currentPage - 1) * PAGE_SIZE + 1;
    const end = Math.min(currentPage * PAGE_SIZE, totalItems);
    
    // Build pagination HTML
    let html = '<div style="display:flex;gap:8px;align-items:center;justify-content:center;margin-top:10px;">';
    
    // Previous button
    html += `<button class="catpicker-page-btn" data-page="${currentPage - 1}" ${currentPage === 1 ? 'disabled' : ''}>Prev</button>`;
    
    // Page indicator
    html += `<span style="font-size:14px;color:#666;">Page ${currentPage} of ${totalPages} (${start}-${end} of ${totalItems})</span>`;
    
    // Next button
    html += `<button class="catpicker-page-btn" data-page="${currentPage + 1}" ${currentPage === totalPages ? 'disabled' : ''}>Next</button>`;
    
    html += '</div>';
    
    pager.innerHTML = html;
    
    // Add click handlers to pagination buttons
    pager.querySelectorAll('button.catpicker-page-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const page = Number(e.currentTarget.getAttribute('data-page'));
        
        // Get currently selected category (to preserve selection across pages)
        const picked = document.querySelector('.catpicker-item[aria-selected="true"]')?.dataset.name;
        
        // Render the requested page
        renderPage(page, filteredCategories, picked);
      });
    });
  }

  // ============================================================================
  // MAIN MODAL FUNCTION (PUBLIC API)
  // ============================================================================
  
  /**
   * Opens the category picker modal
   * This is the main function called from outside this module
   * 
   * @param {Object} options - Configuration object
   * @param {Array} options.categories - Array of category names to choose from
   * @param {string} options.current - Currently selected category
   * @param {Function} options.onChoose - Callback function when user chooses a category
   */
  function openCategoryPicker({categories, current, onChoose}) {
    // Ensure modal exists in the DOM
    ensureModal();
    
    // Get references to modal elements
    const backdrop = document.getElementById('catpickerBackdrop');
    const search = document.getElementById('catpickerSearch');
    const list = document.getElementById('catpickerList');
    const btnUse = document.getElementById('catpickerUse');
    const btnCancel = document.getElementById('catpickerCancel');

    // Store categories and reset to page 1
    allCategories = Array.from(new Set(
      categories.map(c => (c || '').trim()).filter(Boolean)
    ));
    filteredCategories = allCategories;
    currentPage = 1;
    
    // Render initial page
    renderPage(currentPage, filteredCategories, current);

    // ============================================================================
    // SEARCH FUNCTIONALITY
    // ============================================================================
    
    /**
     * Filters categories based on search query
     */
    const filter = () => {
      const q = search.value.toLowerCase().trim();
      
      // Filter categories that include the search query
      filteredCategories = allCategories.filter(c => 
        c.toLowerCase().includes(q)
      );
      
      // Reset to page 1 when search changes
      currentPage = 1;
      
      // Get current selection (if any)
      const picked = document.querySelector('.catpicker-item[aria-selected="true"]')?.dataset.name || current;
      
      // Re-render with filtered list
      renderPage(currentPage, filteredCategories, picked);
    };
    
    // Attach search handler
    search.oninput = filter;

    // ============================================================================
    // MODAL CONTROLS
    // ============================================================================
    
    /**
     * Closes the modal and resets state
     */
    const close = () => {
      backdrop.classList.remove('show');
      search.value = '';
      currentPage = 1;
    };
    
    // Cancel button closes modal without selection
    btnCancel.onclick = close;
    
    // Clicking backdrop (outside dialog) also closes modal
    backdrop.onclick = (e) => {
      if (e.target === backdrop) close();
    };

    // Use button confirms selection
    btnUse.onclick = () => {
      // Get selected category
      const selected = (document.querySelector('.catpicker-item[aria-selected="true"]')?.dataset.name) || current;
      
      // Call the callback function with selected category
      onChoose && onChoose(selected);
      
      // Close modal
      close();
    };

    // Show the modal
    backdrop.classList.add('show');
    
    // Focus search box after a short delay (for better UX)
    setTimeout(() => search.focus(), 50);
  }

  // ============================================================================
  // PUBLIC API
  // ============================================================================
  
  // Expose the main function to the global scope
  // This allows script.js to call: SL_CatPicker.openCategoryPicker(...)
  window.SL_CatPicker = { openCategoryPicker };
})();
