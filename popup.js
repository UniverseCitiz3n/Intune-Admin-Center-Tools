document.addEventListener("DOMContentLoaded", () => {
  // Global state variables
  const state = {
    currentDisplayType: 'config',
    sortDirection: 'asc',
    theme: 'light',
    targetMode: 'device', // New: track whether we're targeting devices or users
    selectedTableRows: new Set(), // Track selected table rows
    dynamicGroups: new Set(), // Track dynamic groups that cannot be modified manually
    lastCheckedGroup: null, // Cache the last group checked via "Check members"
    pagination: {
      currentPage: 1,
      itemsPerPage: 10,
      totalItems: 0,
      totalPages: 0,
      filteredData: [],
      selectedRowIds: new Set() // Track selected rows across pages by unique identifier
    }
  };

  // ── Theme Management Functions ───────────────────────────────────────
  const toggleTheme = () => {
    state.theme = state.theme === 'light' ? 'dark' : 'light';
    applyTheme(state.theme);
    chrome.storage.local.set({ theme: state.theme });
    logMessage(`Theme switched to ${state.theme} mode`);
  };

  const applyTheme = (theme) => {
    document.body.setAttribute('data-theme', theme);
  };

  const initializeTheme = () => {
    chrome.storage.local.get('theme', (data) => {
      if (data.theme) {
        state.theme = data.theme;
        applyTheme(state.theme);
        logMessage(`Theme initialized to ${state.theme} mode`);
      }
    });
  };

  // ── Group Type Management Functions ───────────────────────────────────
  const isDynamicGroup = (groupId) => {
    return state.dynamicGroups.has(groupId);
  };

  const addDynamicGroup = (groupId) => {
    state.dynamicGroups.add(groupId);
  };
  const clearDynamicGroups = () => {
    state.dynamicGroups.clear();
  };

  // fetchGroupWithType: Fetch group details including groupTypes
  const fetchGroupWithType = async (groupId, token) => {
    try {
      const groupData = await fetchJSON(`https://graph.microsoft.com/v1.0/groups/${groupId}?$select=id,displayName,groupTypes`, {
        method: "GET",
        headers: {
          "Authorization": token,
          "Content-Type": "application/json"
        }
      });

      // Check if group is dynamic
      const isDynamic = groupData.groupTypes && groupData.groupTypes.includes('DynamicMembership');
      if (isDynamic) {
        addDynamicGroup(groupData.id);
      }

      return {
        id: groupData.id,
        displayName: groupData.displayName,
        isDynamic: isDynamic
      };
    } catch (error) {
      logMessage(`Error fetching group details for ${groupId}: ${error.message}`);
      return null;
    }
  };

  // ── Utility Functions ───────────────────────────────────────────────
  // filterTable: Filter the table based on input text
  const filterTable = (filterText) => {
    filterText = filterText.toLowerCase();
    if (state.currentDisplayType === 'config') {
      chrome.storage.local.get(['lastConfigAssignments'], (data) => {
        if (data.lastConfigAssignments) {
          const filteredResults = [...data.lastConfigAssignments].filter(item =>
            item.policyName.toLowerCase().includes(filterText)
          );

          updateConfigTable(filteredResults, false); // Pass false to avoid changing currentDisplayType
        }
      });
    } else if (state.currentDisplayType === 'apps') {
      chrome.storage.local.get(['lastAppAssignments'], (data) => {
        if (data.lastAppAssignments) {
          const filteredResults = [...data.lastAppAssignments].filter(item =>
            item.appName.toLowerCase().includes(filterText)
          );

          updateAppTable(filteredResults, false); // Pass false to avoid changing currentDisplayType
        }
      });
    } else if (state.currentDisplayType === 'compliance') {
      chrome.storage.local.get(['lastComplianceAssignments'], (data) => {
        if (data.lastComplianceAssignments) {
          const filteredResults = [...data.lastComplianceAssignments].filter(item =>
            item.policyName.toLowerCase().includes(filterText)
          );

          updateComplianceTable(filteredResults, false); // Pass false to avoid changing currentDisplayType
        }
      });
    } else if (state.currentDisplayType === 'pwsh') {
      chrome.storage.local.get(['lastPwshAssignments'], (data) => {
        if (data.lastPwshAssignments) {
          const filteredResults = [...data.lastPwshAssignments].filter(item =>
            item.scriptName.toLowerCase().includes(filterText)
          );

          updatePwshTable(filteredResults, false); // Pass false to avoid changing currentDisplayType
        }
      });
    } else if (state.currentDisplayType === 'groupMembers') {
      chrome.storage.local.get(['lastGroupMembers'], (data) => {
        if (data.lastGroupMembers) {
          const filteredResults = [...data.lastGroupMembers].filter(item => {
            const name = (item.displayName || '').toLowerCase();
            const upn = (item.userPrincipalName || '').toLowerCase();
            const deviceId = (item.deviceId || '').toLowerCase();
            return name.includes(filterText) || upn.includes(filterText) || deviceId.includes(filterText);
          });

          updateGroupMembersTable(filteredResults, false);
        }
      });
    } else if (state.currentDisplayType === 'groupAssignments') {
      chrome.storage.local.get(['lastGroupAssignments'], (data) => {
        if (data.lastGroupAssignments) {
          const filteredResults = [...data.lastGroupAssignments].filter(item =>
            item.configName.toLowerCase().includes(filterText) ||
            item.configType.toLowerCase().includes(filterText)
          );

          updateGroupAssignmentsTable(filteredResults, false);
        }
      });
    }
  };

  // ── Pagination Functions ────────────────────────────────────────────
  // updatePaginationState: Update pagination state with filtered data
  const updatePaginationState = (data, filterText = '') => {
    const filteredData = filterText ? 
      data.filter(item => {
        const searchText = filterText.toLowerCase();
        if (state.currentDisplayType === 'config') {
          return item.policyName.toLowerCase().includes(searchText);
        } else if (state.currentDisplayType === 'apps') {
          return item.appName.toLowerCase().includes(searchText);
        } else if (state.currentDisplayType === 'compliance') {
          return item.policyName.toLowerCase().includes(searchText);
        } else if (state.currentDisplayType === 'pwsh') {
          return item.scriptName.toLowerCase().includes(searchText);
        } else if (state.currentDisplayType === 'groupMembers') {
          const name = (item.displayName || '').toLowerCase();
          const upn = (item.userPrincipalName || '').toLowerCase();
          const deviceId = (item.deviceId || '').toLowerCase();
          return name.includes(searchText) || upn.includes(searchText) || deviceId.includes(searchText);
        } else if (state.currentDisplayType === 'groupAssignments') {
          return item.configName.toLowerCase().includes(searchText) ||
            item.configType.toLowerCase().includes(searchText);
        }
        return true;
      }) : data;

    state.pagination.filteredData = filteredData;
    state.pagination.totalItems = filteredData.length;
    state.pagination.totalPages = Math.ceil(filteredData.length / state.pagination.itemsPerPage);
    
    // Reset to page 1 if current page is beyond available pages
    if (state.pagination.currentPage > state.pagination.totalPages) {
      state.pagination.currentPage = Math.max(1, state.pagination.totalPages);
    }
  };

  // getCurrentPageData: Get data for current page
  const getCurrentPageData = () => {
    const startIndex = (state.pagination.currentPage - 1) * state.pagination.itemsPerPage;
    const endIndex = startIndex + state.pagination.itemsPerPage;
    return state.pagination.filteredData.slice(startIndex, endIndex);
  };

  // generatePageNumbers: Generate smart page number display
  const generatePageNumbers = () => {
    const { currentPage, totalPages } = state.pagination;
    const pages = [];
    
    if (totalPages <= 7) {
      // Show all pages if 7 or fewer
      for (let i = 1; i <= totalPages; i++) {
        pages.push(i);
      }
    } else {
      // Smart pagination with ellipsis
      if (currentPage <= 4) {
        // Near beginning: 1 2 3 4 5 ... last
        for (let i = 1; i <= 5; i++) {
          pages.push(i);
        }
        pages.push('...');
        pages.push(totalPages);
      } else if (currentPage >= totalPages - 3) {
        // Near end: 1 ... last-4 last-3 last-2 last-1 last
        pages.push(1);
        pages.push('...');
        for (let i = totalPages - 4; i <= totalPages; i++) {
          pages.push(i);
        }
      } else {
        // Middle: 1 ... current-1 current current+1 ... last
        pages.push(1);
        pages.push('...');
        for (let i = currentPage - 1; i <= currentPage + 1; i++) {
          pages.push(i);
        }
        pages.push('...');
        pages.push(totalPages);
      }
    }
    
    return pages;
  };

  // updatePaginationControls: Update pagination UI elements
  const updatePaginationControls = () => {
    const container = document.getElementById('paginationContainer');
    const resultsSpan = document.getElementById('paginationResults');
    const pageNumbers = document.getElementById('pageNumbers');
    const prevBtn = document.getElementById('prevPageBtn');
    const nextBtn = document.getElementById('nextPageBtn');

    if (state.pagination.totalItems === 0) {
      container.style.display = 'none';
      return;
    }

    container.style.display = 'flex';

    // Update results info
    const startItem = (state.pagination.currentPage - 1) * state.pagination.itemsPerPage + 1;
    const endItem = Math.min(state.pagination.currentPage * state.pagination.itemsPerPage, state.pagination.totalItems);
    resultsSpan.textContent = `Showing ${startItem}–${endItem} of ${state.pagination.totalItems} results`;

    // Update page numbers
    const pages = generatePageNumbers();
    pageNumbers.innerHTML = '';
    
    pages.forEach(page => {
      if (page === '...') {
        const ellipsis = document.createElement('span');
        ellipsis.className = 'page-ellipsis';
        ellipsis.textContent = '...';
        pageNumbers.appendChild(ellipsis);
      } else {
        const pageBtn = document.createElement('button');
        pageBtn.className = 'page-number';
        pageBtn.textContent = page;
        if (page === state.pagination.currentPage) {
          pageBtn.classList.add('active');
        }
        pageBtn.addEventListener('click', () => goToPage(page));
        pageNumbers.appendChild(pageBtn);
      }
    });

    // Update prev/next buttons
    prevBtn.disabled = state.pagination.currentPage === 1;
    nextBtn.disabled = state.pagination.currentPage === state.pagination.totalPages;
  };

  // clearTableAndPagination: Helper function to clear table and hide pagination
  const clearTableAndPagination = () => {
    document.getElementById("configTableBody").innerHTML = '';
    document.getElementById('paginationContainer').style.display = 'none';
    state.pagination.totalItems = 0;
    state.pagination.totalPages = 0;
    state.pagination.currentPage = 1;
    state.pagination.filteredData = [];
    state.pagination.selectedRowIds.clear(); // Clear selection when clearing table
  };

  // exportTableToCsv: Export current table data to CSV file
  const exportTableToCsv = () => {
    const data = state.pagination.filteredData;
    if (!data || data.length === 0) {
      showResultNotification('No data to export', 'error');
      return;
    }

    let csvContent = '';
    let headers = [];
    let rows = [];

    // Build CSV based on current display type
    if (state.currentDisplayType === 'config') {
      headers = ['Profile Name', 'Group Name', 'Membership Type', 'Target Type'];
      rows = data.map(item => [
        item.policyName,
        item.targets[0].groupName,
        item.targets[0].membershipType,
        item.targets[0].targetType
      ]);
    } else if (state.currentDisplayType === 'apps') {
      headers = ['App Name', 'Group Name', 'Membership Type', 'Target Type'];
      rows = data.map(item => [
        item.appName,
        item.targets[0].groupName,
        item.targets[0].membershipType,
        item.targets[0].targetType
      ]);
    } else if (state.currentDisplayType === 'compliance') {
      headers = ['Policy Name', 'Group Name', 'Membership Type', 'Target Type'];
      rows = data.map(item => [
        item.policyName,
        item.targets[0].groupName,
        item.targets[0].membershipType,
        item.targets[0].targetType
      ]);
    } else if (state.currentDisplayType === 'pwsh') {
      headers = ['Script Name', 'Group Name', 'Membership Type', 'Target Type'];
      rows = data.map(item => [
        item.scriptName,
        item.targets[0].groupName,
        item.targets[0].membershipType,
        item.targets[0].targetType
      ]);
    } else if (state.currentDisplayType === 'groupMembers') {
      headers = ['Display Name', 'UPN/Device ID', 'Object ID', 'Type'];
      rows = data.map(item => [
        item.displayName || '',
        item.userPrincipalName || item.deviceId || '',
        item.id || '',
        (item['@odata.type'] || '').split('.').pop()
      ]);
    } else if (state.currentDisplayType === 'groupAssignments') {
      headers = ['Configuration Name', 'Configuration Type', 'Intent'];
      rows = data.map(item => [
        item.configName || '',
        item.configType || '',
        item.intent || ''
      ]);
    }

    // Escape CSV values (handle commas, quotes, newlines)
    const escapeCsvValue = (value) => {
      if (value === null || value === undefined) return '';
      const stringValue = String(value);
      if (stringValue.includes(',') || stringValue.includes('"') || stringValue.includes('\n')) {
        return `"${stringValue.replace(/"/g, '""')}"`;
      }
      return stringValue;
    };

    // Build CSV string
    csvContent = headers.map(escapeCsvValue).join(',') + '\n';
    csvContent += rows.map(row => row.map(escapeCsvValue).join(',')).join('\n');

    // Create and download file
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    
    // Generate filename based on display type and timestamp
    const timestamp = new Date().toISOString().slice(0, 10);
    const typeNames = {
      'config': 'configuration_assignments',
      'apps': 'app_assignments',
      'compliance': 'compliance_assignments',
      'pwsh': 'powershell_scripts',
      'groupMembers': 'group_members',
      'groupAssignments': 'group_assignments'
    };
    const filename = `intune_${typeNames[state.currentDisplayType] || 'export'}_${timestamp}.csv`;
    link.setAttribute('download', filename);
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    showResultNotification(`Exported ${data.length} rows to ${filename}`, 'success');
    logMessage(`CSV exported: ${filename} (${data.length} rows)`);
  };

  // generateRowId: Generate unique ID for a table row based on content
  const generateRowId = (rowData) => {
    if (state.currentDisplayType === 'config') {
      return `${rowData.policyName}-${rowData.targets[0].groupName}-${rowData.targets[0].targetType}`;
    } else if (state.currentDisplayType === 'apps') {
      return `${rowData.appName}-${rowData.targets[0].groupName}-${rowData.targets[0].targetType}`;
    } else if (state.currentDisplayType === 'compliance') {
      return `${rowData.policyName}-${rowData.targets[0].groupName}-${rowData.targets[0].targetType}`;
    } else if (state.currentDisplayType === 'pwsh') {
      return `${rowData.scriptName}-${rowData.targets[0].groupName}-${rowData.targets[0].targetType}`;
    } else if (state.currentDisplayType === 'groupMembers') {
      // Generate unique ID based on member's object ID
      return `member-${rowData.id}`;
    }
    return '';
  };

  // restoreRowSelection: Restore row selection after rendering
  const restoreRowSelection = () => {
    document.querySelectorAll('#configTableBody tr').forEach((row) => {
      const rowId = row.dataset.rowId;
      if (rowId && state.pagination.selectedRowIds.has(rowId)) {
        row.classList.add('table-row-selected');
      }
    });
  };

  // clearTableSelection: Clear all table row selections
  const clearTableSelection = () => {
    document.querySelectorAll('.table-row-selected').forEach(row => {
      row.classList.remove('table-row-selected');
    });
    state.pagination.selectedRowIds.clear();
  };

  // goToPage: Navigate to specific page
  const goToPage = (page) => {
    if (page < 1 || page > state.pagination.totalPages) return;
    
    state.pagination.currentPage = page;
    renderCurrentPage();
    updatePaginationControls();
  };

  // renderCurrentPage: Render current page data in table
  const renderCurrentPage = () => {
    const currentPageData = getCurrentPageData();
    
    if (state.currentDisplayType === 'config') {
      renderConfigTablePage(currentPageData);
    } else if (state.currentDisplayType === 'apps') {
      renderAppTablePage(currentPageData);
    } else if (state.currentDisplayType === 'compliance') {
      renderComplianceTablePage(currentPageData);
    } else if (state.currentDisplayType === 'pwsh') {
      renderPwshTablePage(currentPageData);
    } else if (state.currentDisplayType === 'groupMembers') {
      renderGroupMembersTablePage(currentPageData);
    } else if (state.currentDisplayType === 'groupAssignments') {
      renderGroupAssignmentsTablePage(currentPageData);
    }
  };

  // Individual table rendering functions
  const renderConfigTablePage = (assignments) => {
    let rows = '';
    let rowIndex = (state.pagination.currentPage - 1) * state.pagination.itemsPerPage;
    
    assignments.forEach(policy => {
      policy.targets.forEach(target => {
        // Generate unique row ID
        const rowData = { policyName: policy.policyName, targets: [target] };
        const rowId = generateRowId(rowData);
        
        // Check if this is a virtual assignment or dynamic group
        const isVirtualAssignment = target.groupName === 'All Devices' || target.groupName === 'All Users';
        const isDynamicGroupAssignment = target.groupId && isDynamicGroup(target.groupId);
        const isDisabled = isVirtualAssignment || isDynamicGroupAssignment;
        const disabledClass = isDisabled ? 'table-row-disabled' : 'table-row-selectable';

        // Add tooltip for disabled rows
        let tooltipText = '';
        if (isVirtualAssignment) {
          tooltipText = ' title="Virtual group"';
        } else if (isDynamicGroupAssignment) {
          tooltipText = ' title="Dynamic group – cannot modify manually"';
        }

        rows += `<tr class="${disabledClass}" data-row-index="${rowIndex}" data-row-id="${rowId}"${tooltipText}>
          <td style="word-wrap: break-word; white-space: normal;">${policy.policyName}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.groupName}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.membershipType}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.targetType}</td>
        </tr>`;
        rowIndex++;
      });
    });

    document.getElementById("configTableBody").innerHTML = rows;

    // Add event listeners for row clicks
    document.querySelectorAll('#configTableBody tr').forEach((row, index) => {
      row.addEventListener('click', (e) => handleTableRowClick(row, index));
    });

    // Restore selection state
    restoreRowSelection();
  };

  const renderAppTablePage = (assignments) => {
    let rows = '';
    let rowIndex = (state.pagination.currentPage - 1) * state.pagination.itemsPerPage;
    
    assignments.forEach(app => {
      app.targets.forEach(target => {
        // Generate unique row ID
        const rowData = { appName: app.appName, targets: [target] };
        const rowId = generateRowId(rowData);
        
        // Check if this is a virtual assignment or dynamic group
        const isVirtualAssignment = target.groupName === 'All Devices' || target.groupName === 'All Users';
        const isDynamicGroupAssignment = target.groupId && isDynamicGroup(target.groupId);
        const isDisabled = isVirtualAssignment || isDynamicGroupAssignment;
        const disabledClass = isDisabled ? 'table-row-disabled' : 'table-row-selectable';

        // Add tooltip for disabled rows
        let tooltipText = '';
        if (isVirtualAssignment) {
          tooltipText = ' title="Virtual group"';
        } else if (isDynamicGroupAssignment) {
          tooltipText = ' title="Dynamic group – cannot modify manually"';
        }

        rows += `<tr class="${disabledClass}" data-row-index="${rowIndex}" data-row-id="${rowId}"${tooltipText}>
          <td style="word-wrap: break-word; white-space: normal;">${app.appName}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.groupName}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.membershipType}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.targetType}</td>
        </tr>`;
        rowIndex++;
      });
    });

    document.getElementById("configTableBody").innerHTML = rows;

    // Add event listeners for row clicks
    document.querySelectorAll('#configTableBody tr').forEach((row, index) => {
      row.addEventListener('click', (e) => handleTableRowClick(row, index));
    });

    // Restore selection state
    restoreRowSelection();
  };

  const renderComplianceTablePage = (assignments) => {
    let rows = '';
    let rowIndex = (state.pagination.currentPage - 1) * state.pagination.itemsPerPage;
    
    assignments.forEach(policy => {
      policy.targets.forEach(target => {
        // Generate unique row ID
        const rowData = { policyName: policy.policyName, targets: [target] };
        const rowId = generateRowId(rowData);
        
        // Check if this is a virtual assignment or dynamic group
        const isVirtualAssignment = target.groupName === 'All Devices' || target.groupName === 'All Users';
        const isDynamicGroupAssignment = target.groupId && isDynamicGroup(target.groupId);
        const isDisabled = isVirtualAssignment || isDynamicGroupAssignment;
        const disabledClass = isDisabled ? 'table-row-disabled' : 'table-row-selectable';

        // Add tooltip for disabled rows
        let tooltipText = '';
        if (isVirtualAssignment) {
          tooltipText = ' title="Virtual group"';
        } else if (isDynamicGroupAssignment) {
          tooltipText = ' title="Dynamic group – cannot modify manually"';
        }

        rows += `<tr class="${disabledClass}" data-row-index="${rowIndex}" data-row-id="${rowId}"${tooltipText}>
          <td style="word-wrap: break-word; white-space: normal;">${policy.policyName}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.groupName}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.membershipType}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.targetType}</td>
        </tr>`;
        rowIndex++;
      });
    });

    document.getElementById("configTableBody").innerHTML = rows;

    // Add event listeners for row clicks
    document.querySelectorAll('#configTableBody tr').forEach((row, index) => {
      row.addEventListener('click', (e) => handleTableRowClick(row, index));
    });

    // Restore selection state
    restoreRowSelection();
  };

  const renderPwshTablePage = (assignments) => {
    let rows = '';
    let rowIndex = (state.pagination.currentPage - 1) * state.pagination.itemsPerPage;
    
    assignments.forEach(script => {
      script.targets.forEach(target => {
        // Generate unique row ID
        const rowData = { scriptName: script.scriptName, targets: [target] };
        const rowId = generateRowId(rowData);
        
        // Check if this is a virtual assignment or dynamic group
        const isVirtualAssignment = target.groupName === 'All Devices' || target.groupName === 'All Users';
        const isDynamicGroupAssignment = target.groupId && isDynamicGroup(target.groupId);
        const isDisabled = isVirtualAssignment || isDynamicGroupAssignment;
        const disabledClass = isDisabled ? 'table-row-disabled' : 'table-row-selectable';

        // Add tooltip for disabled rows
        let tooltipText = '';
        if (isVirtualAssignment) {
          tooltipText = ' title="Virtual group"';
        } else if (isDynamicGroupAssignment) {
          tooltipText = ' title="Dynamic group – cannot modify manually"';
        }

        rows += `<tr class="${disabledClass}" data-row-index="${rowIndex}" data-row-id="${rowId}"${tooltipText}>
          <td style="word-wrap: break-word; white-space: normal;">${script.scriptName}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.groupName}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.membershipType}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.targetType}</td>
        </tr>`;
        rowIndex++;
      });
    });

    document.getElementById("configTableBody").innerHTML = rows;

    // Add event listeners for row clicks
    document.querySelectorAll('#configTableBody tr').forEach((row, index) => {
      row.addEventListener('click', (e) => handleTableRowClick(row, index));
    });

    // Restore selection state
    restoreRowSelection();
  };

  const renderGroupMembersTablePage = (members) => {
    let rows = '';
    let rowIndex = (state.pagination.currentPage - 1) * state.pagination.itemsPerPage;

    members.forEach(member => {
      const objectType = (member['@odata.type'] || '').split('.').pop();
      const upnOrId = member.userPrincipalName || member.deviceId || '';
      const objectId = member.id || '';
      
      // Generate unique row ID
      const rowData = { id: objectId };
      const rowId = generateRowId(rowData);
      
      // Add selectable class for row selection
      const selectableClass = 'table-row-selectable';
      
      // For devices, show device ID in the UPN/Device ID column and object ID separately
      // For users, show UPN in the UPN/Device ID column and object ID separately
      rows += `<tr class="${selectableClass}" data-row-index="${rowIndex}" data-row-id="${rowId}">
        <td style="word-wrap: break-word; white-space: normal;">${member.displayName || ''}</td>
        <td style="word-wrap: break-word; white-space: normal;">${upnOrId}</td>
        <td style="word-wrap: break-word; white-space: normal;">${objectId}</td>
        <td style="word-wrap: break-word; white-space: normal;">${objectType}</td>
      </tr>`;
      rowIndex++;
    });

    document.getElementById("configTableBody").innerHTML = rows;
    
    // Add event listeners for row clicks
    document.querySelectorAll('#configTableBody tr').forEach((row, index) => {
      row.addEventListener('click', (e) => handleTableRowClick(row, index));
    });

    // Restore selection state
    restoreRowSelection();
  };

  const renderGroupAssignmentsTablePage = (assignments) => {
    let rows = '';
    let rowIndex = (state.pagination.currentPage - 1) * state.pagination.itemsPerPage;

    assignments.forEach(assignment => {
      rows += `<tr data-row-index="${rowIndex}">
        <td style="word-wrap: break-word; white-space: normal;">${assignment.configName || ''}</td>
        <td style="word-wrap: break-word; white-space: normal;">${assignment.configType || ''}</td>
        <td style="word-wrap: break-word; white-space: normal;">${assignment.intent || ''}</td>
      </tr>`;
      rowIndex++;
    });

    document.getElementById("configTableBody").innerHTML = rows;
  };

  // fetchJSON: Helper to fetch and parse JSON responses
  const fetchJSON = async (url, options = {}) => {
    const response = await fetch(url, options);
    return parseJSON(response);
  };

  // verifyMdmUrl: Verify the current tab URL and extract mdmDeviceId
  const verifyMdmUrl = async () => {
    return new Promise((resolve, reject) => {
      chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
        if (!tabs || !tabs[0]) {
          const error = 'No active tab found.';
          logMessage(error);
          showResultNotification(error, 'error');
          return reject(new Error(error));
        }
        const url = tabs[0].url;
        const sanitizedUrl = url.replace(/\/[\w-]+(?=\/|$)/g, '/[REDACTED]');
        logMessage(`Active tab URL: ${sanitizedUrl}`);
        const mdmMatch = url.match(/mdmDeviceId\/([\w-]+)/i);
        if (!mdmMatch) {
          const error = 'mdmDeviceId not found in URL.';
          logMessage(error);
          showResultNotification(error, 'error');
          return reject(new Error(error));
        }
        resolve({ mdmDeviceId: mdmMatch[1] });
      });
    });
  };

  // getToken: Retrieve token from Chrome storage
  const getToken = async () => {
    return new Promise((resolve, reject) => {
      chrome.storage.local.get("msGraphToken", (data) => {
        if (data.msGraphToken) {
          resolve(data.msGraphToken);
        } else {
          const error = 'No token captured. Please login first.';
          logMessage(error);
          showResultNotification(error, 'error');
          reject(new Error(error));
        }
      });
    });  
  };  // getAllGroupsMap: Get groups map (device & user) for lookups, with dynamic group tracking
  const getAllGroupsMap = async (deviceObjectId, userObjectId, token) => {
    const headers = {
      "Authorization": token,
      "Content-Type": "application/json",
      "ConsistencyLevel": "eventual"
    };

    // Helper function to fetch all pages of results with pagination support
    const fetchAllPages = async (baseUrl) => {
      let allResults = [];
      let url = baseUrl;
      let pageCount = 0;
      const maxPages = 50; // Safety limit to prevent infinite loops
      
      try {
        while (url && pageCount < maxPages) {
          pageCount++;
          logMessage(`getAllGroupsMap: Fetching page ${pageCount} from ${url.substring(0, 100)}...`);
          
          const result = await fetchJSON(url, { method: "GET", headers });
          if (result.value && Array.isArray(result.value)) {
            allResults = allResults.concat(result.value);
            logMessage(`getAllGroupsMap: Page ${pageCount} returned ${result.value.length} groups`);
          }
          
          // Check for next page
          url = result['@odata.nextLink'] || null;
        }
        
        if (pageCount >= maxPages && url) {
          logMessage(`getAllGroupsMap: Warning - reached maximum page limit (${maxPages}), some groups may be missing`);
        }
        
        logMessage(`getAllGroupsMap: Fetched total of ${allResults.length} groups across ${pageCount} pages`);
        return { value: allResults };
      } catch (error) {
        logMessage(`getAllGroupsMap: Error fetching from ${baseUrl}: ${error.message}`);
        return { value: [] };
      }
    };

    // Separate endpoints for direct and transitive memberships with pagination support
    const directEndpoints = [
      `https://graph.microsoft.com/beta/devices/${deviceObjectId}/memberOf?$select=id,displayName,groupTypes&$orderBy=displayName%20asc&$top=999&$count=true`
    ];
    
    const transitiveEndpoints = [
      `https://graph.microsoft.com/beta/devices/${deviceObjectId}/transitiveMemberOf?$select=id,displayName,groupTypes&$orderBy=displayName%20asc&$top=999&$count=true`
    ];

    if (userObjectId) {
      directEndpoints.push(
        `https://graph.microsoft.com/beta/users/${userObjectId}/memberOf?$select=id,displayName,groupTypes&$orderBy=displayName%20asc&$top=999&$count=true`
      );
      transitiveEndpoints.push(
        `https://graph.microsoft.com/beta/users/${userObjectId}/transitiveMemberOf?$select=id,displayName,groupTypes&$orderBy=displayName%20asc&$top=999&$count=true`
      );
    }

    const [directResults, transitiveResults] = await Promise.all([
      Promise.all(directEndpoints.map(url => fetchAllPages(url))),
      Promise.all(transitiveEndpoints.map(url => fetchAllPages(url)))
    ]);

    const directGroupMap = new Map();
    const transitiveGroupMap = new Map();
    const allGroupsMap = new Map();

    // Process direct group memberships
    directResults.forEach(result => {
      (result.value || []).forEach(group => {
        if (group['@odata.type'] === '#microsoft.graph.group') {
          directGroupMap.set(group.id, group.displayName);
          allGroupsMap.set(group.id, group.displayName);

          // Track dynamic groups
          const isDynamic = group.groupTypes && group.groupTypes.includes('DynamicMembership');
          if (isDynamic) {
            addDynamicGroup(group.id);
          }
        }
      });
    });

    // Process transitive group memberships
    transitiveResults.forEach(result => {
      (result.value || []).forEach(group => {
        if (group['@odata.type'] === '#microsoft.graph.group') {
          transitiveGroupMap.set(group.id, group.displayName);
          allGroupsMap.set(group.id, group.displayName);

          // Track dynamic groups
          const isDynamic = group.groupTypes && group.groupTypes.includes('DynamicMembership');
          if (isDynamic) {
            addDynamicGroup(group.id);
          }
        }
      });
    });

    logMessage(`getAllGroupsMap: Final results - Direct: ${directGroupMap.size}, Transitive: ${transitiveGroupMap.size}, Total: ${allGroupsMap.size} groups`);

    return {
      allGroups: allGroupsMap,
      directGroups: directGroupMap,
      transitiveGroups: transitiveGroupMap
    };
  };

  // resolveGroupInfo: Attempt to resolve group name and check membership for groups not in groupMaps
  const resolveGroupInfo = async (groupId, deviceObjectId, userObjectId, token, groupMaps) => {
    try {
      logMessage(`resolveGroupInfo: Attempting to resolve group ${groupId}`);
      
      // First, try to get group basic info
      const groupData = await fetchJSON(`https://graph.microsoft.com/beta/groups/${groupId}?$select=id,displayName,groupTypes`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });
      
      if (!groupData || !groupData.displayName) {
        logMessage(`resolveGroupInfo: Could not resolve group name for ${groupId}`);
        return null;
      }
      
      const groupName = groupData.displayName;
      const isDynamic = groupData.groupTypes && groupData.groupTypes.includes('DynamicMembership');
      
      // Now check if the device or user is actually a member of this group
      const membershipChecks = [];
      
      // Check device membership (both direct and transitive)
      membershipChecks.push(
        checkGroupMembership(groupId, deviceObjectId, 'device', token)
      );
      
      // Check user membership if user exists
      if (userObjectId) {
        membershipChecks.push(
          checkGroupMembership(groupId, userObjectId, 'user', token)
        );
      }
      
      const membershipResults = await Promise.all(membershipChecks);
      const deviceMembership = membershipResults[0];
      const userMembership = userObjectId ? membershipResults[1] : { isDirect: false, isTransitive: false };
      
      // If neither device nor user is a member, return null (maintain core logic)
      if (!deviceMembership.isDirect && !deviceMembership.isTransitive && 
          !userMembership.isDirect && !userMembership.isTransitive) {
        logMessage(`resolveGroupInfo: Device/User is not a member of group ${groupName} (${groupId}), correctly filtering out`);
        return null;
      }
      
      // Determine membership type
      let membershipType = 'Direct';
      if ((deviceMembership.isDirect || userMembership.isDirect)) {
        membershipType = 'Direct';
        logMessage(`resolveGroupInfo: Found DIRECT membership for group ${groupName} (${groupId})`);
      } else if ((deviceMembership.isTransitive || userMembership.isTransitive)) {
        membershipType = 'Transitive';
        logMessage(`resolveGroupInfo: Found TRANSITIVE membership for group ${groupName} (${groupId})`);
      }
      
      // Update groupMaps for future use
      groupMaps.allGroups.set(groupId, groupName);
      if (membershipType === 'Direct') {
        groupMaps.directGroups.set(groupId, groupName);
      } else {
        groupMaps.transitiveGroups.set(groupId, groupName);
      }
      
      // Track dynamic groups
      if (isDynamic) {
        addDynamicGroup(groupId);
      }
      
      logMessage(`resolveGroupInfo: Successfully resolved missing group ${groupName} (${groupId}) with ${membershipType} membership`);
      return { groupName, membershipType, isDynamic };
      
    } catch (error) {
      logMessage(`resolveGroupInfo: Error resolving group ${groupId}: ${error.message}`);
      return null;
    }
  };

  // checkGroupMembership: Check if an object is a member of a specific group
  const checkGroupMembership = async (groupId, objectId, objectType, token) => {
    const headers = {
      "Authorization": token,
      "Content-Type": "application/json",
      "ConsistencyLevel": "eventual"
    };
    
    try {
      const baseUrl = objectType === 'device' 
        ? `https://graph.microsoft.com/beta/devices/${objectId}`
        : `https://graph.microsoft.com/beta/users/${objectId}`;
      
      // Check direct membership
      const directUrl = `${baseUrl}/memberOf?$filter=id eq '${groupId}'&$select=id`;
      const directResult = await fetchJSON(directUrl, { method: "GET", headers });
      const isDirect = directResult.value && directResult.value.length > 0;
      
      // Check transitive membership if not direct
      let isTransitive = false;
      if (!isDirect) {
        const transitiveUrl = `${baseUrl}/transitiveMemberOf?$filter=id eq '${groupId}'&$select=id`;
        const transitiveResult = await fetchJSON(transitiveUrl, { method: "GET", headers });
        isTransitive = transitiveResult.value && transitiveResult.value.length > 0;
      }
      
      return { isDirect, isTransitive };
    } catch (error) {
      logMessage(`checkGroupMembership: Error checking membership for ${objectType} ${objectId} in group ${groupId}: ${error.message}`);
      return { isDirect: false, isTransitive: false };
    }
  };

  // getDirectoryObjectId: Get directory object ID for device or user based on current mode
  const getDirectoryObjectId = async (mdmDeviceId, token, mode = state.targetMode) => {
    logMessage(`getDirectoryObjectId: Getting ${mode} directory object ID`);

    // First, get the managed device data
    const managedDeviceData = await fetchJSON(`https://graph.microsoft.com/beta/deviceManagement/manageddevices('${mdmDeviceId}')`, {
      method: "GET",
      headers: { "Authorization": token, "Content-Type": "application/json" }
    });

    if (mode === 'device') {
      const azureADDeviceId = managedDeviceData.azureADDeviceId;
      if (!azureADDeviceId) throw new Error("azureADDeviceId not found in managed device data.");

      const filter = encodeURIComponent(`deviceId eq '${azureADDeviceId}'`);
      const deviceData = await fetchJSON(`https://graph.microsoft.com/v1.0/devices?$top=100&$filter=${filter}`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });

      if (!deviceData.value || deviceData.value.length === 0) {
        throw new Error("No device found matching the azureADDeviceId.");
      }

      return {
        directoryId: deviceData.value[0].id,
        displayName: deviceData.value[0].displayName || 'Unknown Device'
      };
    } else if (mode === 'user') {
      const userPrincipalName = managedDeviceData.userPrincipalName;
      if (!userPrincipalName || userPrincipalName === 'Unknown user') {
        throw new Error("No user associated with this device.");
      }

      const userData = await fetchJSON(`https://graph.microsoft.com/beta/users?$filter=userPrincipalName eq '${encodeURIComponent(userPrincipalName)}'`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });

      if (!userData.value || userData.value.length === 0) {
        throw new Error("No user found matching the userPrincipalName.");
      }

      return {
        directoryId: userData.value[0].id,
        displayName: userData.value[0].displayName || userPrincipalName
      };
    }

    throw new Error(`Invalid mode: ${mode}`);
  };  // ── UI Update Functions ──────────────────────────────────────────────  
  // updateDeviceNameDisplay: Update the device name display next to "Configuration Assignments"
  const updateDeviceNameDisplay = (deviceData) => {
    const deviceNameElement = document.getElementById('deviceNameDisplay');
    if (deviceNameElement && deviceData && deviceData.deviceName) {
      deviceNameElement.textContent = `- ${deviceData.deviceName}`;
      logMessage(`updateDeviceNameDisplay: Updated device name display to "${deviceData.deviceName}"`);
    }
  };
  // Table row selection helpers
  // (clearTableSelection moved to pagination section)

  // Clear checkbox selections
  const clearCheckboxSelection = () => {
    document.querySelectorAll("#groupResults input[type=checkbox]:checked").forEach(cb => {
      cb.checked = false;
    });
    // Update stored search results to reflect unchecked state
    chrome.storage.local.get(['lastSearchResults'], (data) => {
      if (data.lastSearchResults) {
        const updated = data.lastSearchResults.map(group => ({ ...group, checked: false }));
        chrome.storage.local.set({ lastSearchResults: updated });
      }
    });
    updateActionButtonsState();
  };

  const updateActionButtonsState = () => {
    const selected = document.querySelectorAll('#groupResults input[type=checkbox]:checked');
    const hasDynamic = Array.from(selected).some(cb => isDynamicGroup(cb.value));
    
    // Collect the names of dynamic groups that are selected
    const dynamicGroupNames = Array.from(selected)
      .filter(cb => isDynamicGroup(cb.value))
      .map(cb => cb.dataset.groupName)
      .filter(name => name); // Remove any undefined names
    
    ['addToGroups', 'removeFromGroups'].forEach(id => {
      const btn = document.getElementById(id);
      if (!btn) return;
      if (hasDynamic) {
        btn.classList.add('disabled');
        // Add tooltip showing which dynamic groups are causing the disablement
        const tooltipText = dynamicGroupNames.length > 0 
          ? `Cannot modify dynamic group${dynamicGroupNames.length > 1 ? 's' : ''}: ${dynamicGroupNames.join(', ')}`
          : 'Cannot modify dynamic groups';
        btn.title = tooltipText;
      } else {
        btn.classList.remove('disabled');
        // Remove tooltip when buttons are enabled
        btn.removeAttribute('title');
      }
    });
    
    // Update Clear Members button state
    updateClearMembersButtonState();
  };

  const getSelectedGroupNames = () => {
    const selectedGroups = [];
    document.querySelectorAll('.table-row-selected').forEach(row => {
      let groupNameCell;

      // Determine which column contains the group name based on current table type
      if (state.currentDisplayType === 'pwsh') {
        groupNameCell = row.children[1]; // Group is the 2nd column (index 1) for PowerShell scripts
      } else {
        groupNameCell = row.children[1]; // Group Name is the 2nd column (index 1) for other tables after removing checkbox column
      }

      const groupName = groupNameCell.textContent.trim();

      // Only include actual group names, not virtual assignments
      if (groupName && groupName !== 'All Devices' && groupName !== 'All Users') {
        selectedGroups.push(groupName);
      }
    });
    return [...new Set(selectedGroups)]; // Remove duplicates
  };
  const getAllSelectedGroups = () => {
    // Get groups from search results (exclude dynamic groups from operations)
    const searchResultGroups = [];
    const searchCheckboxes = document.querySelectorAll("#groupResults input[type=checkbox]:checked");
    searchCheckboxes.forEach(cb => {
      // Only include non-dynamic groups
      if (!isDynamicGroup(cb.value)) {
        searchResultGroups.push({
          id: cb.value,
          name: cb.dataset.groupName
        });
      }
    });

    // Get group names from selected table rows (new functionality)
    // Note: Dynamic groups are already filtered out by disabling row selection in table rendering
    const tableSelectedGroupNames = getSelectedGroupNames();

    return {
      searchResults: searchResultGroups,
      tableSelections: tableSelectedGroupNames,
      hasAnySelection: searchResultGroups.length > 0 || tableSelectedGroupNames.length > 0
    };
  }; const handleTableRowClick = (row, rowIndex) => {
    // Don't allow selection of disabled rows
    if (row.classList.contains('table-row-disabled')) {
      return;
    }

    // Clear checkbox selections when selecting table rows
    clearCheckboxSelection();

    const rowId = row.dataset.rowId;
    if (!rowId) return; // Skip if no row ID

    if (row.classList.contains('table-row-selected')) {
      // Deselect the row
      state.pagination.selectedRowIds.delete(rowId);
      row.classList.remove('table-row-selected');
    } else {
      // Select the row
      state.pagination.selectedRowIds.add(rowId);
      row.classList.add('table-row-selected');
    }
  };

  // updateButtonText: Update button text based on current target mode
  const updateButtonText = () => {
    const targetType = state.targetMode === 'device' ? 'Device' : 'User';
    document.getElementById('addBtnText').textContent = `Add`;
    document.getElementById('removeBtnText').textContent = `Remove`;
    logMessage(`updateButtonText: Updated buttons for ${targetType} mode`);
  };

  // handleTargetModeToggle: Handle switching between device and user modes
  const handleTargetModeToggle = (mode) => {
    logMessage(`handleTargetModeToggle: Called with mode '${mode}', current state is '${state.targetMode}'`);
    
    if (state.targetMode === mode) {
      logMessage(`handleTargetModeToggle: No change needed - already in ${mode} mode`);
      return; // No change needed
    }

    state.targetMode = mode;
    logMessage(`handleTargetModeToggle: State updated to '${mode}'`);

    // Update UI
    document.querySelectorAll('.target-type-toggle button').forEach(btn => {
      btn.classList.remove('active');
    });
    document.querySelector(`[data-mode="${mode}"]`).classList.add('active');
    logMessage(`handleTargetModeToggle: UI updated - '${mode}' button is now active`);

    updateButtonText();

    // Save to storage
    chrome.storage.local.set({ targetMode: state.targetMode });

    logMessage(`handleTargetModeToggle: Switched to ${mode} mode`);
  };  // updateTableHeaders: Update table header based on content type
  const updateTableHeaders = (type) => {
    const headerRow = document.querySelector('thead tr');
    let headerContent = '';
    if (type === 'config') {
      headerContent = `
        <th class="sortable" style="word-wrap: break-word; white-space: normal;">Profile Name</th>
        <th style="word-wrap: break-word; white-space: normal;">Group</th>
        <th style="word-wrap: break-word; white-space: normal;">Membership Type</th>
        <th style="word-wrap: break-word; white-space: normal;">Target Type</th>
      `;
    } else if (type === 'apps') {
      headerContent = `
        <th class="sortable" style="word-wrap: break-word; white-space: normal;">App Name</th>
        <th style="word-wrap: break-word; white-space: normal;">Group</th>
        <th style="word-wrap: break-word; white-space: normal;">Membership Type</th>
        <th style="word-wrap: break-word; white-space: normal;">Target Type</th>
      `;
    } else if (type === 'compliance') {
      headerContent = `
        <th class="sortable" style="word-wrap: break-word; white-space: normal;">Policy Name</th>
        <th style="word-wrap: break-word; white-space: normal;">Group</th>
        <th style="word-wrap: break-word; white-space: normal;">Membership Type</th>
        <th style="word-wrap: break-word; white-space: normal;">Target Type</th>
      `;
    } else if (type === 'pwsh') {
      headerContent = `
        <th class="sortable" style="word-wrap: break-word; white-space: normal;">Script Name</th>
        <th style="word-wrap: break-word; white-space: normal;">Group</th>
        <th style="word-wrap: break-word; white-space: normal;">Membership Type</th>
        <th style="word-wrap: break-word; white-space: normal;">Target Type</th>
      `;
    } else if (type === 'groupMembers') {
      headerContent = `
        <th class="sortable" style="word-wrap: break-word; white-space: normal;">Display Name</th>
        <th style="word-wrap: break-word; white-space: normal;">UPN / Device ID</th>
        <th style="word-wrap: break-word; white-space: normal;">Object ID</th>
        <th style="word-wrap: break-word; white-space: normal;">Object Type</th>
      `;
    } else if (type === 'groupAssignments') {
      headerContent = `
        <th class="sortable" style="word-wrap: break-word; white-space: normal;">Configuration Name</th>
        <th style="word-wrap: break-word; white-space: normal;">Configuration Type</th>
        <th style="word-wrap: break-word; white-space: normal;">Intent</th>
      `;
    }
    headerRow.innerHTML = headerContent;
    const sortableHeader = document.querySelector('th.sortable');
    if (sortableHeader) {
      // Apply current sort direction to the header
      sortableHeader.classList.remove('desc', 'asc');
      sortableHeader.classList.add(state.sortDirection);
    }
  };  // updateConfigTable: Update configuration assignments table
  const updateConfigTable = (assignments, updateDisplay = true) => {
    if (updateDisplay) {
      state.currentDisplayType = 'config';
      chrome.storage.local.set({ currentDisplayType: state.currentDisplayType });
    }
    state.pagination.itemsPerPage = 10;
    updateTableHeaders('config');
    assignments.sort((a, b) => a.policyName.localeCompare(b.policyName));
    if (state.sortDirection === 'desc') assignments.reverse();

    // Flatten assignments for pagination (each target becomes a row)
    const flattenedData = [];
    assignments.forEach(policy => {
      policy.targets.forEach(target => {
        flattenedData.push({
          policyName: policy.policyName,
          targets: [target] // Keep single target for rendering
        });
      });
    });

    // Update pagination state
    const filterValue = document.getElementById('profileFilterInput').value.toLowerCase();
    updatePaginationState(flattenedData, filterValue);
    
    // Render current page
    renderCurrentPage();
    updatePaginationControls();

    const sortableHeader = document.querySelector('th.sortable');
    if (sortableHeader) {
      sortableHeader.classList.remove('desc', 'asc');
      sortableHeader.classList.add(state.sortDirection);
    }
  };

  const updateGroupMembersTable = (members, updateDisplay = true) => {
    if (updateDisplay) {
      state.currentDisplayType = 'groupMembers';
      chrome.storage.local.set({ currentDisplayType: state.currentDisplayType });
    }
    state.pagination.itemsPerPage = 10;
    updateTableHeaders('groupMembers');
    members.sort((a, b) => (a.displayName || '').localeCompare(b.displayName || ''));
    if (state.sortDirection === 'desc') members.reverse();

    const flattenedData = members.map(m => ({
      displayName: m.displayName || '',
      userPrincipalName: m.userPrincipalName || '',
      deviceId: m.deviceId || '',
      id: m.id || '', // Include object ID
      ['@odata.type']: m['@odata.type'] || ''
    }));

    updatePaginationState(flattenedData);

    renderCurrentPage();
    updatePaginationControls();

    const sortableHeader = document.querySelector('th.sortable');
    if (sortableHeader) {
      sortableHeader.classList.remove('desc', 'asc');
      sortableHeader.classList.add(state.sortDirection);
    }
  };

  const updateGroupAssignmentsTable = (assignments, updateDisplay = true) => {
    if (updateDisplay) {
      state.currentDisplayType = 'groupAssignments';
      chrome.storage.local.set({ currentDisplayType: state.currentDisplayType });
    }
    state.pagination.itemsPerPage = 10;
    updateTableHeaders('groupAssignments');
    assignments.sort((a, b) => (a.configName || '').localeCompare(b.configName || ''));
    if (state.sortDirection === 'desc') assignments.reverse();

    updatePaginationState(assignments);

    renderCurrentPage();
    updatePaginationControls();

    const sortableHeader = document.querySelector('th.sortable');
    if (sortableHeader) {
      sortableHeader.classList.remove('desc', 'asc');
      sortableHeader.classList.add(state.sortDirection);
    }
  };
  // updateAppTable: Update app assignments table (similar in structure)
  const updateAppTable = (assignments, updateDisplay = true) => {
    if (updateDisplay) {
      state.currentDisplayType = 'apps';
      chrome.storage.local.set({ currentDisplayType: state.currentDisplayType });
    }
    state.pagination.itemsPerPage = 10;
    updateTableHeaders('apps');
    assignments.sort((a, b) => a.appName.localeCompare(b.appName));
    if (state.sortDirection === 'desc') assignments.reverse();

    // Flatten assignments for pagination (each target becomes a row)
    const flattenedData = [];
    assignments.forEach(app => {
      const targets = app.targets.filter(t => t.membershipType !== 'Not Member');
      targets.forEach(target => {
        flattenedData.push({
          appName: app.appName,
          appVersion: app.appVersion,
          installState: app.installState,
          targets: [target] // Keep single target for rendering
        });
      });
    });

    // Update pagination state
    const filterValue = document.getElementById('profileFilterInput').value.toLowerCase();
    updatePaginationState(flattenedData, filterValue);
    
    // Render current page
    renderCurrentPage();
    updatePaginationControls();

    const sortableHeader = document.querySelector('th.sortable');
    if (sortableHeader) {
      sortableHeader.classList.remove('desc', 'asc');
      sortableHeader.classList.add(state.sortDirection);
    }
  };
  // updateComplianceTable: Update compliance assignments table
  const updateComplianceTable = (assignments, updateDisplay = true) => {
    if (updateDisplay) {
      state.currentDisplayType = 'compliance';
      chrome.storage.local.set({ currentDisplayType: state.currentDisplayType });
    }
    state.pagination.itemsPerPage = 10;
    updateTableHeaders('compliance');
    assignments.sort((a, b) => a.policyName.localeCompare(b.policyName));
    if (state.sortDirection === 'desc') assignments.reverse();

    // Flatten assignments for pagination (each target becomes a row)
    const flattenedData = [];
    assignments.forEach(policy => {
      policy.targets.forEach(target => {
        flattenedData.push({
          policyName: policy.policyName,
          complianceState: policy.complianceState,
          targets: [target] // Keep single target for rendering
        });
      });
    });

    // Update pagination state
    const filterValue = document.getElementById('profileFilterInput').value.toLowerCase();
    updatePaginationState(flattenedData, filterValue);
    
    // Render current page
    renderCurrentPage();
    updatePaginationControls();

    const sortableHeader = document.querySelector('th.sortable');
    if (sortableHeader) {
      sortableHeader.classList.remove('desc', 'asc');
      sortableHeader.classList.add(state.sortDirection);
    }
  };
  // updatePwshTable: Update PowerShell scripts table
  const updatePwshTable = (assignments, updateDisplay = true) => {
    if (updateDisplay) {
      state.currentDisplayType = 'pwsh';
      chrome.storage.local.set({ currentDisplayType: state.currentDisplayType });
    }
    state.pagination.itemsPerPage = 10;
    updateTableHeaders('pwsh');
    assignments.sort((a, b) => a.scriptName.localeCompare(b.scriptName));
    if (state.sortDirection === 'desc') assignments.reverse();

    // Transform flattened data into grouped structure like config/compliance assignments
    const scriptMap = new Map();
    assignments.forEach(item => {
      if (!scriptMap.has(item.scriptName)) {
        scriptMap.set(item.scriptName, {
          scriptName: item.scriptName,
          description: item.description || '',
          targets: []
        });
      }
      
      // Use the stored target type and determine membership type
      let targetType = item.targetType || 'Device'; // Default to Device if not specified
      let membershipType = 'Assigned';
      
      if (item.targetName === 'All Devices' || item.targetName === 'All Users') {
        membershipType = 'Direct';
      }
      
      scriptMap.get(item.scriptName).targets.push({
        groupName: item.targetName,
        groupId: item.targetGroupId,
        targetType: targetType,
        membershipType: membershipType
      });
    });

    const groupedAssignments = Array.from(scriptMap.values());

    // Flatten assignments for pagination (each target becomes a row)
    const flattenedData = [];
    groupedAssignments.forEach(script => {
      script.targets.forEach(target => {
        flattenedData.push({
          scriptName: script.scriptName,
          description: script.description,
          targets: [target] // Keep single target for rendering
        });
      });
    });

    // Update pagination state
    const filterValue = document.getElementById('profileFilterInput').value.toLowerCase();
    updatePaginationState(flattenedData, filterValue);
    
    // Render current page
    renderCurrentPage();
    updatePaginationControls();

    const sortableHeader = document.querySelector('th.sortable');
    if (sortableHeader) {
      sortableHeader.classList.remove('desc', 'asc');
      sortableHeader.classList.add(state.sortDirection);
    }
  };
  // ── State Restoration Functions ─────────────────────────────────────────
  const restoreFilterValue = () => {
    chrome.storage.local.get(
      ['profileFilterValue', 'currentDisplayType', 'targetMode', 'lastComplianceAssignments', 'lastAppAssignments', 'lastConfigAssignments', 'lastPwshAssignments', 'lastGroupMembers', 'lastGroupAssignments'],
      (data) => {
        // Restore target mode
        if (data.targetMode) {
          logMessage(`restoreFilterValue: Restoring target mode to '${data.targetMode}' from storage`);
          // Don't set state.targetMode first - let handleTargetModeToggle do it
          handleTargetModeToggle(data.targetMode);
        } else {
          logMessage(`restoreFilterValue: No stored target mode, using default '${state.targetMode}'`);
          // Ensure UI reflects the default state on first load
          handleTargetModeToggle(state.targetMode);
        }

        if (data.currentDisplayType) {
          state.currentDisplayType = data.currentDisplayType;
          clearTableAndPagination();
          if (state.currentDisplayType === 'compliance' && data.lastComplianceAssignments) {
            chrome.storage.local.remove(['lastConfigAssignments', 'lastAppAssignments', 'lastPwshAssignments']);
            updateComplianceTable(data.lastComplianceAssignments, false);
          } else if (state.currentDisplayType === 'apps' && data.lastAppAssignments) {
            chrome.storage.local.remove(['lastConfigAssignments', 'lastComplianceAssignments', 'lastPwshAssignments']);
            updateAppTable(data.lastAppAssignments, false);
          } else if (state.currentDisplayType === 'config' && data.lastConfigAssignments) {
            chrome.storage.local.remove(['lastAppAssignments', 'lastComplianceAssignments', 'lastPwshAssignments']);
            updateConfigTable(data.lastConfigAssignments, false);
          } else if (state.currentDisplayType === 'pwsh' && data.lastPwshAssignments) {
            chrome.storage.local.remove(['lastConfigAssignments', 'lastAppAssignments', 'lastComplianceAssignments', 'lastGroupMembers']);
            updatePwshTable(data.lastPwshAssignments, false);
          } else if (state.currentDisplayType === 'groupMembers' && data.lastGroupMembers) {
            chrome.storage.local.remove(['lastConfigAssignments', 'lastAppAssignments', 'lastComplianceAssignments', 'lastPwshAssignments']);
            updateGroupMembersTable(data.lastGroupMembers, false);
          } else if (state.currentDisplayType === 'groupAssignments' && data.lastGroupAssignments) {
            chrome.storage.local.remove(['lastConfigAssignments', 'lastAppAssignments', 'lastComplianceAssignments', 'lastPwshAssignments', 'lastGroupMembers']);
            updateGroupAssignmentsTable(data.lastGroupAssignments, false);
          }
        } else {
          clearTableAndPagination();
        }
        if (data.profileFilterValue) {
          document.getElementById('profileFilterInput').value = data.profileFilterValue;
          filterTable(data.profileFilterValue);
        }
      }
    );
  };
  const restoreState = () => {
    chrome.storage.local.get(['lastSearchResults', 'lastSearchQuery'], (data) => {
      if (data.lastSearchQuery) {
        document.getElementById("groupSearchInput").value = data.lastSearchQuery;
      }
      if (data.lastSearchResults) {
        const resultsDiv = document.getElementById("groupResults");
        resultsDiv.innerHTML = "";

        // Clear and rebuild dynamic groups tracking from stored results
        clearDynamicGroups();

        data.lastSearchResults.forEach(group => {
          // Restore dynamic group tracking
          if (group.isDynamic) {
            addDynamicGroup(group.id);
          }

          const item = document.createElement("p");
          const label = document.createElement("label");
          const checkbox = document.createElement("input");

          checkbox.type = "checkbox";
          checkbox.value = group.id;
          checkbox.id = "group-" + group.id;
          checkbox.className = "filled-in";
          checkbox.dataset.groupName = group.displayName;
          if (group.checked) checkbox.checked = true;

          if (group.isDynamic) {
            checkbox.title = "Dynamic group – cannot modify manually";
          }
          const span = document.createElement("span");
          span.textContent = group.displayName;
          if (group.isDynamic) {
            span.title = "Dynamic group – cannot modify manually";
          }


          label.appendChild(checkbox);
          label.appendChild(span);
          item.appendChild(label);
          resultsDiv.appendChild(item);
        });
        updateActionButtonsState();
      }
    });
  };
  // ── Event Handler Functions ─────────────────────────────────────────────
  // Handle Group Search
  const handleSearchGroup = async () => {
    logMessage("searchGroup clicked");
    const query = document.getElementById("groupSearchInput").value.trim();
    if (!query) {
      logMessage("searchGroup: No query entered");
      showResultNotification("Enter group name to search.", "info");
      return;
    }
    try {
      showProcessingNotification(`Searching for groups containing "${query}"...`);
      
      const token = await getToken();
      logMessage("searchGroup: Token found, proceeding with fetch");

      // Fetch groups with groupTypes information
      const groupsData = await fetchJSON(`https://graph.microsoft.com/v1.0/groups?$search="displayName:${query}"&$select=id,displayName,groupTypes&$top=10`, {
        method: "GET",
        headers: {
          "Authorization": token,
          "Content-Type": "application/json",
          "ConsistencyLevel": "eventual"
        }
      });

      logMessage("searchGroup: Received groups data");
      const resultsDiv = document.getElementById("groupResults");
      resultsDiv.innerHTML = "";

      if (!groupsData.value || groupsData.value.length === 0) {
        logMessage("searchGroup: No groups found");
        resultsDiv.textContent = "No groups found.";
        showResultNotification(`No groups found matching "${query}".`, "info");
        return;
      }

      // Clear previous dynamic groups tracking for search results
      clearDynamicGroups();

      const searchResults = groupsData.value.map(g => {
        const isDynamic = g.groupTypes && g.groupTypes.includes('DynamicMembership');
        if (isDynamic) {
          addDynamicGroup(g.id);
        }

        return {
          id: g.id,
          displayName: g.displayName,
          isDynamic: isDynamic,
          checked: false
        };
      });

      searchResults.forEach(group => {
        const item = document.createElement("p");
        const label = document.createElement("label");
        const checkbox = document.createElement("input");

        checkbox.type = "checkbox";
        checkbox.value = group.id;
        checkbox.id = "group-" + group.id;
        checkbox.className = "filled-in";
        checkbox.dataset.groupName = group.displayName;

        if (group.isDynamic) {
          checkbox.title = "Dynamic group – cannot modify manually";
        }

        const span = document.createElement("span");
        span.textContent = group.displayName;
        if (group.isDynamic) {
          span.title = "Dynamic group – cannot modify manually";
        }

        label.appendChild(checkbox);
        label.appendChild(span);
        item.appendChild(label); resultsDiv.appendChild(item);
      });
      updateActionButtonsState();

      chrome.storage.local.set({ lastSearchResults: searchResults, lastSearchQuery: query });
      showResultNotification(`Found ${searchResults.length} group(s) matching "${query}".`, "success");
    } catch (error) {
      logMessage(`searchGroup: Error - ${error.message}`);
      showResultNotification('Error: ' + error.message, 'error');
    }
  };

  // Handle Adding Device/User to Selected Groups
  const handleAddToGroups = async () => {
    const targetType = state.targetMode === 'device' ? 'device' : 'user';
    logMessage(`addToGroups clicked (${targetType} mode)`);

    if (document.getElementById('addToGroups').classList.contains('disabled')) {
      showResultNotification('Cannot modify dynamic groups.', 'error');
      return;
    }

    const allSelected = getAllSelectedGroups();
    if (!allSelected.hasAnySelection) {
      logMessage("addToGroups: No groups selected");
      alert("Select at least one group from search results or table rows.");
      return;
    }

    // Show processing notification with count
    const totalCount = allSelected.searchResults.length + allSelected.tableSelections.length;
    showProcessingNotification(`Adding ${targetType} to ${totalCount} group(s)...`);

    try {
      const { mdmDeviceId } = await verifyMdmUrl();
      const token = await getToken();
      logMessage(`addToGroups: Extracted mdmDeviceId ${mdmDeviceId}`);

      const { directoryId, displayName } = await getDirectoryObjectId(mdmDeviceId, token, state.targetMode);
      logMessage(`addToGroups: Got directory ID for ${targetType}: ${directoryId}`);

      const promises = [];

      // Process search result groups (existing logic with group IDs)
      allSelected.searchResults.forEach(group => {
        promises.push(addToSingleGroup(group.id, directoryId, token, group.name));
      });

      // Process table selection groups (need to resolve names to IDs)
      for (const groupName of allSelected.tableSelections) {
        try {
          const groupData = await fetchJSON(`https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent(groupName)}'&$select=id,displayName`, {
            method: "GET",
            headers: { "Authorization": token, "Content-Type": "application/json" }
          });

          if (groupData.value && groupData.value.length > 0) {
            const groupId = groupData.value[0].id;
            promises.push(addToSingleGroup(groupId, directoryId, token, groupName));
          } else {
            promises.push(Promise.resolve({ groupName, error: "Group not found" }));
          }
        } catch (e) {
          promises.push(Promise.resolve({ groupName, error: e.message }));
        }
      }

      const results = await Promise.all(promises);
      let success = true;
      let message = '';
      results.forEach(r => {
        if (r.error) {
          success = false;
          const identifier = r.groupId || r.groupName || 'Unknown Group';
          message += `Group ${identifier}: Error - ${r.error}\n`;
        } else {
          const identifier = r.groupId || r.groupName || 'Unknown Group';
          message += `Group ${identifier}: Success\n`;
        }
      });

      const successMsg = `Successfully added ${targetType} "${displayName}" to groups`;
      const errorMsg = `Some groups could not be added for ${targetType} "${displayName}"\n${message}`;
      showResultNotification(success ? successMsg : errorMsg, success ? 'success' : 'error');
    } catch (error) {
      logMessage(`addToGroups: Error - ${error.message}`);
      showResultNotification(`Failed to add ${targetType} to groups: ${error.message}`, 'error');
    }
  };

  // Helper function to add to a single group
  const addToSingleGroup = async (groupId, directoryId, token, groupName = null) => {
    const postUrl = `https://graph.microsoft.com/beta/groups/${groupId}/members/$ref`;
    const body = JSON.stringify({
      "@odata.id": `https://graph.microsoft.com/beta/directoryObjects/${directoryId}`
    });

    try {
      await fetchJSON(postUrl, {
        method: "POST",
        headers: { "Authorization": token, "Content-Type": "application/json" },
        body: body
      });
      return { groupId, groupName, result: "Success" };
    } catch (e) {
      return { groupId, groupName, error: e.message };
    }
  };  // Handle Removing Device/User from Selected Groups
  const handleRemoveFromGroups = async () => {
    const targetType = state.targetMode === 'device' ? 'device' : 'user';
    logMessage(`removeFromGroups clicked (${targetType} mode)`);

    if (document.getElementById('removeFromGroups').classList.contains('disabled')) {
      showResultNotification('Cannot modify dynamic groups.', 'error');
      return;
    }

    const allSelected = getAllSelectedGroups();
    if (!allSelected.hasAnySelection) {
      logMessage("removeFromGroups: No groups selected");
      alert("Select at least one group from search results or table rows.");
      return;
    }

    // Show processing notification with count
    const totalCount = allSelected.searchResults.length + allSelected.tableSelections.length;
    showProcessingNotification(`Removing ${targetType} from ${totalCount} group(s)...`);

    try {
      const { mdmDeviceId } = await verifyMdmUrl();
      const token = await getToken();
      logMessage(`removeFromGroups: Extracted mdmDeviceId ${mdmDeviceId}`);

      const { directoryId, displayName } = await getDirectoryObjectId(mdmDeviceId, token, state.targetMode);
      logMessage(`removeFromGroups: Got directory ID for ${targetType}: ${directoryId}`);

      const promises = [];

      // Process search result groups (existing logic with group IDs)
      allSelected.searchResults.forEach(group => {
        promises.push(removeFromSingleGroup(group.id, directoryId, token, group.name));
      });

      // Process table selection groups (need to resolve names to IDs)
      for (const groupName of allSelected.tableSelections) {
        try {
          const groupData = await fetchJSON(`https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent(groupName)}'&$select=id,displayName`, {
            method: "GET",
            headers: { "Authorization": token, "Content-Type": "application/json" }
          });

          if (groupData.value && groupData.value.length > 0) {
            const groupId = groupData.value[0].id;
            promises.push(removeFromSingleGroup(groupId, directoryId, token, groupName));
          } else {
            promises.push(Promise.resolve({ groupName, error: "Group not found" }));
          }
        } catch (e) {
          promises.push(Promise.resolve({ groupName, error: e.message }));
        }
      }

      const results = await Promise.all(promises);
      let success = true;
      let message = '';
      results.forEach(r => {
        if (r.error) {
          success = false;
          const identifier = r.groupId || r.groupName || 'Unknown Group';
          message += `Group ${identifier}: Error - ${r.error}\n`;
        } else {
          const identifier = r.groupId || r.groupName || 'Unknown Group';
          message += `Group ${identifier}: Success\n`;
        }
      });

      const successMsg = `Successfully removed ${targetType} "${displayName}" from groups`;
      const errorMsg = `Some groups could not be removed for ${targetType} "${displayName}"\n${message}`;
      showResultNotification(success ? successMsg : errorMsg, success ? 'success' : 'error');
    } catch (error) {
      logMessage(`removeFromGroups: Error - ${error.message}`);
      showResultNotification(`Failed to remove ${targetType} from groups: ${error.message}`, 'error');
    }
  };

  // Helper function to remove from a single group
  const removeFromSingleGroup = async (groupId, directoryId, token, groupName = null) => {
    const deleteUrl = `https://graph.microsoft.com/v1.0/groups/${groupId}/members/${directoryId}/$ref`;

    try {
      const response = await fetch(deleteUrl, {
        method: "DELETE",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });
      if (response.ok) return { groupId, groupName, result: "Removed" };
      const text = await response.text();
      return { groupId, groupName, error: text || response.statusText };
    } catch (e) {
      return { groupId, groupName, error: e.message };
    }
  };

  const fetchAllGroupMembers = async (groupId, token) => {
    const headers = {
      "Authorization": token,
      "Content-Type": "application/json",
      "ConsistencyLevel": "eventual"  // Required for advanced query capabilities
    };
    // Note: @odata.type cannot be used in $select - it's always returned by default
    const selectFields = 'id,displayName,userType,appId,mail,onPremisesSyncEnabled,deviceId,userPrincipalName';

    logMessage(`fetchAllGroupMembers: Fetching members for group ${groupId}`);

    // First, verify we can access the group itself and get detailed group info
    let groupInfo = null;
    try {
      const groupInfoUrl = `https://graph.microsoft.com/beta/groups/${groupId}?$select=id,displayName,groupTypes,visibility,mailEnabled,securityEnabled,membershipRule,membershipRuleProcessingState`;
      groupInfo = await fetchJSON(groupInfoUrl, { method: 'GET', headers });
      logMessage(`fetchAllGroupMembers: Successfully accessed group info: ${JSON.stringify(groupInfo)}`);
    } catch (error) {
      logMessage(`fetchAllGroupMembers: Failed to access group info - ${error.message}`);
      throw new Error(`Cannot access group. ${error.message}`);
    }

    // Using beta endpoint to avoid known issue #25984 where v1.0 doesn't return service principals
    const directUrl = `https://graph.microsoft.com/beta/groups/${groupId}/members?$select=${selectFields}&$top=999&$count=true`;
    const transitiveUrl = `https://graph.microsoft.com/beta/groups/${groupId}/transitiveMembers?$select=${selectFields}&$top=999&$count=true`;

    try {
      logMessage(`fetchAllGroupMembers: Making API calls to:\n- Direct: ${directUrl}\n- Transitive: ${transitiveUrl}`);
      
      const [directRes, transitiveRes] = await Promise.all([
        fetchJSON(directUrl, { method: 'GET', headers }),
        fetchJSON(transitiveUrl, { method: 'GET', headers })
      ]);

      // Log the full response structure for debugging
      logMessage(`fetchAllGroupMembers: Direct members full response: ${JSON.stringify(directRes)}`);
      logMessage(`fetchAllGroupMembers: Transitive members full response: ${JSON.stringify(transitiveRes)}`);

      // Check for errors in response
      if (directRes.error || transitiveRes.error) {
        logMessage(`fetchAllGroupMembers: API returned errors - falling back to basic request`);
        throw new Error(`API Error: ${directRes.error?.message || transitiveRes.error?.message}`);
      }

      logMessage(`fetchAllGroupMembers: Direct members response analysis: ${JSON.stringify({
        hasValue: !!directRes.value,
        valueType: typeof directRes.value,
        isArray: Array.isArray(directRes.value),
        count: directRes.value?.length || 0,
        odataCount: directRes['@odata.count'],
        odataContext: directRes['@odata.context'],
        keys: Object.keys(directRes || {})
      })}`);
      
      logMessage(`fetchAllGroupMembers: Transitive members response analysis: ${JSON.stringify({
        hasValue: !!transitiveRes.value,
        valueType: typeof transitiveRes.value,
        isArray: Array.isArray(transitiveRes.value),
        count: transitiveRes.value?.length || 0,
        odataCount: transitiveRes['@odata.count'],
        odataContext: transitiveRes['@odata.context'],
        keys: Object.keys(transitiveRes || {})
      })}`);

      const directMembers = directRes.value || [];
      const transitiveMembers = transitiveRes.value || [];
      
      // Combine both direct and transitive members
      const combined = [...directMembers, ...transitiveMembers];
      const unique = [];
      const seen = new Set();
      
      combined.forEach(m => {
        // Include all member types including nested groups
        if (!seen.has(m.id)) {
          seen.add(m.id);
          unique.push(m);
          logMessage(`fetchAllGroupMembers: Adding member: ${m.displayName || 'Unknown'} (${m['@odata.type'] || 'Unknown type'})`);
        }
      });

      // Get the exact count from API responses (use @odata.count if available, otherwise use unique.length)
      const exactCount = directRes['@odata.count'] || transitiveRes['@odata.count'] || unique.length;
      
      // Determine if group is dynamic
      const isDynamic = groupInfo?.groupTypes?.includes('DynamicMembership') || false;
      
      logMessage(`fetchAllGroupMembers: Final result - ${unique.length} unique members from ${combined.length} total entries, exact count: ${exactCount}, isDynamic: ${isDynamic}`);
      
      return { 
        members: unique, 
        totalCount: unique.length,
        exactCount: exactCount,
        groupInfo: {
          isDynamic: isDynamic,
          membershipRule: groupInfo?.membershipRule || null,
          membershipRuleProcessingState: groupInfo?.membershipRuleProcessingState || null,
          groupTypes: groupInfo?.groupTypes || []
        }
      };
      
    } catch (error) {
      logMessage(`fetchAllGroupMembers: Error fetching group members - ${error.message}`);
      
      // Try fallback approach with just direct members if transitive fails
      try {
        logMessage(`fetchAllGroupMembers: Trying fallback - direct members only`);
        const fallbackRes = await fetchJSON(directUrl, { method: 'GET', headers });
        logMessage(`fetchAllGroupMembers: Fallback response: ${JSON.stringify(fallbackRes)}`);
        const fallbackMembers = fallbackRes.value || [];
        
        // Include all member types including groups
        const unique = fallbackMembers;
        const exactCount = fallbackRes['@odata.count'] || unique.length;
        const isDynamic = groupInfo?.groupTypes?.includes('DynamicMembership') || false;
        
        logMessage(`fetchAllGroupMembers: Fallback successful - ${unique.length} direct members`);
        
        return { 
          members: unique, 
          totalCount: unique.length,
          exactCount: exactCount,
          groupInfo: {
            isDynamic: isDynamic,
            membershipRule: groupInfo?.membershipRule || null,
            membershipRuleProcessingState: groupInfo?.membershipRuleProcessingState || null,
            groupTypes: groupInfo?.groupTypes || []
          }
        };
      } catch (fallbackError) {
        logMessage(`fetchAllGroupMembers: Fallback also failed - ${fallbackError.message}`);
        // Continue to basic fallback approaches
      }
    }

    // If we get here, try multiple basic approaches
    logMessage(`fetchAllGroupMembers: Trying basic fallback approaches`);
    
    // Try 1: Basic request without advanced query parameters
    try {
      const basicUrl = `https://graph.microsoft.com/beta/groups/${groupId}/members?$select=${selectFields}`;
      const basicHeaders = {
        "Authorization": token,
        "Content-Type": "application/json"
      };
      const basicRes = await fetchJSON(basicUrl, { method: 'GET', headers: basicHeaders });
      logMessage(`fetchAllGroupMembers: Basic request response: ${JSON.stringify(basicRes)}`);
      
      if (basicRes.value && Array.isArray(basicRes.value)) {
        // Include all member types including groups
        const basicMembers = basicRes.value;
        logMessage(`fetchAllGroupMembers: Basic request returned ${basicMembers.length} members`);
        
        if (basicMembers.length > 0) {
          const exactCount = basicRes['@odata.count'] || basicMembers.length;
          const isDynamic = groupInfo?.groupTypes?.includes('DynamicMembership') || false;
          return { 
            members: basicMembers, 
            totalCount: basicMembers.length,
            exactCount: exactCount,
            groupInfo: {
              isDynamic: isDynamic,
              membershipRule: groupInfo?.membershipRule || null,
              membershipRuleProcessingState: groupInfo?.membershipRuleProcessingState || null,
              groupTypes: groupInfo?.groupTypes || []
            }
          };
        }
      }
    } catch (basicError) {
      logMessage(`fetchAllGroupMembers: Basic request failed: ${basicError.message}`);
    }
    
    // Try 2: Check if it's a security group with different approach
    try {
      const expandUrl = `https://graph.microsoft.com/beta/groups/${groupId}?$expand=members($select=${selectFields})`;
      const expandHeaders = {
        "Authorization": token,
        "Content-Type": "application/json"
      };
      const expandRes = await fetchJSON(expandUrl, { method: 'GET', headers: expandHeaders });
      logMessage(`fetchAllGroupMembers: Expand request response: ${JSON.stringify(expandRes)}`);
      
      if (expandRes.members && Array.isArray(expandRes.members)) {
        // Include all member types including groups
        const expandMembers = expandRes.members;
        logMessage(`fetchAllGroupMembers: Expand request returned ${expandMembers.length} members`);
        
        if (expandMembers.length > 0) {
          const isDynamic = groupInfo?.groupTypes?.includes('DynamicMembership') || false;
          return { 
            members: expandMembers, 
            totalCount: expandMembers.length,
            exactCount: expandMembers.length,
            groupInfo: {
              isDynamic: isDynamic,
              membershipRule: groupInfo?.membershipRule || null,
              membershipRuleProcessingState: groupInfo?.membershipRuleProcessingState || null,
              groupTypes: groupInfo?.groupTypes || []
            }
          };
        }
      }
    } catch (expandError) {
      logMessage(`fetchAllGroupMembers: Expand request failed: ${expandError.message}`);
    }
    
    // Try 3: Check for count only to see if there are members but we can't see them
    try {
      const countUrl = `https://graph.microsoft.com/beta/groups/${groupId}/members/$count`;
      const countHeaders = {
        "Authorization": token,
        "Content-Type": "application/json",
        "ConsistencyLevel": "eventual"
      };
      const countRes = await fetch(countUrl, { method: 'GET', headers: countHeaders });
      const countText = await countRes.text();
      logMessage(`fetchAllGroupMembers: Count request returned: ${countText}`);
      
      const memberCount = parseInt(countText, 10);
      if (memberCount > 0) {
        logMessage(`fetchAllGroupMembers: Group has ${memberCount} members but they're not visible - likely permission issue`);
        throw new Error(`Group has ${memberCount} members but you don't have permission to view them. This may be a security group with hidden membership.`);
      }
    } catch (countError) {
      logMessage(`fetchAllGroupMembers: Count request failed: ${countError.message}`);
    }

    // If we get here, the group truly appears to have no members
    const isDynamic = groupInfo?.groupTypes?.includes('DynamicMembership') || false;
    return { 
      members: [], 
      totalCount: 0,
      exactCount: 0,
      groupInfo: {
        isDynamic: isDynamic,
        membershipRule: groupInfo?.membershipRule || null,
        membershipRuleProcessingState: groupInfo?.membershipRuleProcessingState || null,
        groupTypes: groupInfo?.groupTypes || []
      }
    };
  };

  // Handle Checking Group Members
  const handleCheckGroupMembers = async () => {
    logMessage("checkGroupMembers clicked");
    const selected = document.querySelectorAll("#groupResults input[type=checkbox]:checked");
    if (selected.length !== 1) {
      showResultNotification('Select exactly one group.', 'error');
      return;
    }

    document.getElementById('profileFilterInput').value = '';
    chrome.storage.local.set({ profileFilterValue: '' });

    clearTableSelection();

    const groupId = selected[0].value;
    const groupName = selected[0].dataset.groupName;

    logMessage(`checkGroupMembers: Selected group - ID: ${groupId}, Name: ${groupName}`);

    try {
      const token = await getToken();
      logMessage("checkGroupMembers: Token retrieved successfully");

      showProcessingNotification(`Fetching members for group "${groupName}"...`);

      const { members, totalCount, exactCount, groupInfo, note } = await fetchAllGroupMembers(groupId, token);

      logMessage(`checkGroupMembers: Retrieved ${totalCount} members for group ${groupName}, exact count: ${exactCount}, isDynamic: ${groupInfo.isDynamic}`);

      // Cache the group info for use by other features (e.g., Clear members)
      state.lastCheckedGroup = {
        groupId: groupId,
        groupName: groupName,
        isDynamic: groupInfo.isDynamic,
        membershipRule: groupInfo.membershipRule
      };

      // Clear other data types from storage
      chrome.storage.local.remove(['lastConfigAssignments','lastAppAssignments','lastComplianceAssignments','lastPwshAssignments']);
      chrome.storage.local.set({ lastGroupMembers: members });

      // Update UI - show exact count and group type
      const groupType = groupInfo.isDynamic ? 'Dynamic' : 'Assigned';
      let displayText = `- ${groupName} (${exactCount} members, ${groupType})`;
      if (note) {
        displayText += ` - ${note}`;
      }
      document.getElementById('deviceNameDisplay').textContent = displayText;

      // Show/hide dynamic query section
      const dynamicQuerySection = document.getElementById('dynamicQuerySection');
      if (dynamicQuerySection) {
        if (groupInfo.isDynamic && groupInfo.membershipRule) {
          dynamicQuerySection.style.display = 'block';
          document.getElementById('dynamicQueryContent').textContent = groupInfo.membershipRule;
        } else {
          dynamicQuerySection.style.display = 'none';
        }
      }

      updateGroupMembersTable(members);

      if (totalCount === 0) {
        if (note) {
          showResultNotification(`Group "${groupName}" loaded. ${note}`, 'info');
        } else {
          showResultNotification(`Group "${groupName}" has no members or you don't have permission to view them.`, 'warning');
        }
      } else {
        showResultNotification(`Successfully loaded ${totalCount} members for group "${groupName}".`, 'success');
      }
      
    } catch (error) {
      logMessage(`checkGroupMembers: Error - ${error.message}`);
      
      let errorMessage = 'Failed to load group members: ' + error.message;
      
      // Provide more specific error messages for common issues
      if (error.message.includes('403') || error.message.includes('Forbidden')) {
        errorMessage = `Access denied. You don't have permission to view members of group "${groupName}". Contact your administrator.`;
      } else if (error.message.includes('404') || error.message.includes('Not Found')) {
        errorMessage = `Group "${groupName}" was not found or has been deleted.`;
      } else if (error.message.includes('401') || error.message.includes('Unauthorized')) {
        errorMessage = 'Authentication failed. Please refresh the page and try again.';
      }
      
      showResultNotification(errorMessage, 'error');
    }
  };

  // ══════════════════════════════════════════════════════════════
  // Clear Group Members Feature
  // ══════════════════════════════════════════════════════════════

  // State for clear members modal
  const clearMembersState = {
    selectedGroupId: null,
    selectedGroupName: null,
    isGroupDynamic: false,
    allMembers: [],
    actionType: null, // 'selected' or 'all'
    scope: {
      users: true,
      devices: true,
      nestedGroups: false
    }
  };

  // Show the clear members modal
  const showClearMembersModal = () => {
    const modal = document.getElementById('clearMembersModal');
    modal.style.display = 'flex';
    
    // Reset modal state
    document.getElementById('confirmationSection').style.display = 'none';
    document.getElementById('progressSection').style.display = 'none';
    document.getElementById('resultsSection').style.display = 'none';
    document.querySelector('.clear-action-buttons').style.display = 'flex';
    document.querySelector('.modal-description').style.display = 'block';
    
    // Reset scope checkboxes
    document.getElementById('scopeUsers').checked = true;
    document.getElementById('scopeDevices').checked = true;
    document.getElementById('scopeNestedGroups').checked = false;
    
    // Reset typed confirm input
    document.getElementById('typedConfirmInput').value = '';
    document.getElementById('typedConfirmSection').style.display = 'none';
  };

  // Hide the clear members modal and reset state
  const hideClearMembersModal = () => {
    const modal = document.getElementById('clearMembersModal');
    modal.style.display = 'none';
    
    // Reset modal state
    clearMembersState.selectedGroupId = null;
    clearMembersState.selectedGroupName = null;
    clearMembersState.isGroupDynamic = false;
    clearMembersState.allMembers = [];
    clearMembersState.actionType = null;
    clearMembersState.scope = {
      users: true,
      devices: true,
      nestedGroups: false
    };
  };

  // Update member counts in modal
  const updateMemberCounts = (selectedCount, totalCount) => {
    const selectedBtn = document.getElementById('clearSelectedMembersBtn');
    const allBtn = document.getElementById('clearAllMembersBtn');
    
    if (selectedCount > 0) {
      selectedBtn.disabled = false;
      document.getElementById('selectedMembersCount').textContent = 
        `${selectedCount} member${selectedCount !== 1 ? 's' : ''} selected`;
    } else {
      selectedBtn.disabled = true;
      document.getElementById('selectedMembersCount').textContent = 'No rows selected';
    }
    
    document.getElementById('allMembersCount').textContent = 
      `${totalCount} total member${totalCount !== 1 ? 's' : ''}`;
  };

  // Get selected members from table
  const getSelectedMembers = () => {
    const selectedIds = Array.from(state.pagination.selectedRowIds);
    return clearMembersState.allMembers.filter(member => 
      selectedIds.includes(member.id)
    );
  };

  // Filter members by scope
  const filterMembersByScope = (members) => {
    return members.filter(member => {
      const odataType = member['@odata.type'] || '';
      
      if (odataType.includes('#microsoft.graph.user')) {
        return clearMembersState.scope.users;
      } else if (odataType.includes('#microsoft.graph.device')) {
        return clearMembersState.scope.devices;
      } else if (odataType.includes('#microsoft.graph.group')) {
        return clearMembersState.scope.nestedGroups;
      }
      
      // Default: include if type is unknown and users are enabled
      return clearMembersState.scope.users;
    });
  };

  // Show confirmation section
  const showConfirmationSection = (actionType) => {
    clearMembersState.actionType = actionType;
    
    // Hide action buttons
    document.querySelector('.clear-action-buttons').style.display = 'none';
    document.querySelector('.modal-description').style.display = 'none';
    
    // Show confirmation section
    document.getElementById('confirmationSection').style.display = 'block';
    
    // Update confirmation message (escape HTML to prevent XSS)
    const escapeHtml = (text) => {
      const div = document.createElement('div');
      div.textContent = text;
      return div.innerHTML;
    };
    
    const escapedGroupName = escapeHtml(clearMembersState.selectedGroupName);
    let count, message;
    if (actionType === 'selected') {
      const selectedMembers = getSelectedMembers();
      count = selectedMembers.length;
      message = `You are about to remove <strong>${count} selected member${count !== 1 ? 's' : ''}</strong> from the group "${escapedGroupName}".`;
      document.getElementById('typedConfirmSection').style.display = 'none';
      document.getElementById('confirmRemovalBtn').disabled = false;
    } else {
      count = clearMembersState.allMembers.length;
      message = `You are about to remove <strong>ALL ${count} member${count !== 1 ? 's' : ''}</strong> from the group "${escapedGroupName}". This action cannot be undone.`;
      document.getElementById('typedConfirmSection').style.display = 'block';
      document.getElementById('confirmRemovalBtn').disabled = true;
    }
    
    document.getElementById('confirmationMessage').innerHTML = message;
  };

  // Validate typed confirmation
  const validateTypedConfirmation = () => {
    if (clearMembersState.actionType !== 'all') {
      return true;
    }
    
    const input = document.getElementById('typedConfirmInput').value;
    return input === 'REMOVE ALL';
  };

  // Remove members from group using Graph API
  const removeMembersFromGroup = async (groupId, memberIds, token) => {
    const results = {
      total: memberIds.length,
      removed: 0,
      failed: 0,
      failures: []
    };

    // Process in batches of 20 (Graph API batch limit)
    const batchSize = 20;
    
    for (let i = 0; i < memberIds.length; i += batchSize) {
      const batch = memberIds.slice(i, i + batchSize);
      
      // Update progress
      const progress = Math.min(100, Math.round((i / memberIds.length) * 100));
      document.getElementById('progressBar').style.width = `${progress}%`;
      document.getElementById('progressDetails').textContent = 
        `Processing ${Math.min(i + batchSize, memberIds.length)} of ${memberIds.length}...`;
      
      // Try batch request first
      try {
        // Create a mapping of batch request IDs to member IDs
        const idMapping = {};
        const batchRequests = batch.map((memberId, index) => {
          const requestId = `${i + index}`;
          idMapping[requestId] = memberId;
          return {
            id: requestId,
            method: 'DELETE',
            url: `/groups/${groupId}/members/${memberId}/$ref`
          };
        });

        const batchResponse = await fetchJSON('https://graph.microsoft.com/v1.0/$batch', {
          method: 'POST',
          headers: {
            'Authorization': token,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({ requests: batchRequests })
        });

        // Process batch responses
        if (batchResponse.responses) {
          for (const response of batchResponse.responses) {
            if (response.status === 204 || response.status === 200) {
              results.removed++;
            } else {
              results.failed++;
              const memberId = idMapping[response.id];
              const member = clearMembersState.allMembers.find(m => m.id === memberId);
              results.failures.push({
                memberId: memberId || 'Unknown',
                memberName: member ? (member.displayName || member.userPrincipalName || 'Unknown') : 'Unknown',
                reason: response.body?.error?.message || `HTTP ${response.status}`
              });
            }
          }
        }
      } catch (batchError) {
        logMessage(`Batch removal failed, falling back to individual requests: ${batchError.message}`);
        
        // Fallback: individual DELETE requests with retry logic
        for (const memberId of batch) {
          let retries = 0;
          let success = false;
          
          while (retries < 3 && !success) {
            try {
              const deleteUrl = `https://graph.microsoft.com/v1.0/groups/${groupId}/members/${memberId}/$ref`;
              const response = await fetch(deleteUrl, {
                method: 'DELETE',
                headers: {
                  'Authorization': token,
                  'Content-Type': 'application/json'
                }
              });

              if (response.status === 204 || response.status === 200) {
                results.removed++;
                success = true;
              } else if (response.status === 429 || response.status === 503) {
                // Throttling - wait and retry
                const retryAfter = parseInt(response.headers.get('Retry-After')) || (retries + 1) * 2;
                logMessage(`Throttled (${response.status}), waiting ${retryAfter}s before retry ${retries + 1}/3`);
                await new Promise(resolve => setTimeout(resolve, retryAfter * 1000));
                retries++;
              } else {
                // Other error
                results.failed++;
                const errorBody = await response.json().catch(() => ({}));
                const member = clearMembersState.allMembers.find(m => m.id === memberId);
                results.failures.push({
                  memberId: memberId,
                  memberName: member ? (member.displayName || member.userPrincipalName || 'Unknown') : 'Unknown',
                  reason: errorBody?.error?.message || `HTTP ${response.status}`
                });
                break;
              }
            } catch (error) {
              retries++;
              if (retries >= 3) {
                results.failed++;
                const member = clearMembersState.allMembers.find(m => m.id === memberId);
                results.failures.push({
                  memberId: memberId,
                  memberName: member ? (member.displayName || member.userPrincipalName || 'Unknown') : 'Unknown',
                  reason: error.message
                });
                break;
              }
              retries++;
              await new Promise(resolve => setTimeout(resolve, retries * 1000));
            }
          }
        }
      }
    }

    // Final progress update
    document.getElementById('progressBar').style.width = '100%';
    
    return results;
  };

  // Execute the removal
  const executeRemoval = async () => {
    // Hide confirmation, show progress
    document.getElementById('confirmationSection').style.display = 'none';
    document.getElementById('progressSection').style.display = 'block';
    
    // Update scope from checkboxes
    clearMembersState.scope.users = document.getElementById('scopeUsers').checked;
    clearMembersState.scope.devices = document.getElementById('scopeDevices').checked;
    clearMembersState.scope.nestedGroups = document.getElementById('scopeNestedGroups').checked;
    
    // Get members to remove
    let membersToRemove;
    if (clearMembersState.actionType === 'selected') {
      membersToRemove = getSelectedMembers();
    } else {
      membersToRemove = clearMembersState.allMembers;
    }
    
    // Filter by scope
    membersToRemove = filterMembersByScope(membersToRemove);
    
    if (membersToRemove.length === 0) {
      showResultsSection({
        total: 0,
        removed: 0,
        failed: 0,
        failures: []
      });
      return;
    }
    
    // Get member IDs
    const memberIds = membersToRemove.map(m => m.id);
    
    document.getElementById('progressMessage').textContent = 
      `Removing ${memberIds.length} member${memberIds.length !== 1 ? 's' : ''}...`;
    document.getElementById('progressBar').style.width = '0%';
    document.getElementById('progressDetails').textContent = 'Initializing...';
    
    try {
      const token = await getToken();
      const results = await removeMembersFromGroup(
        clearMembersState.selectedGroupId,
        memberIds,
        token
      );
      
      showResultsSection(results);
      
      // If we removed members from the currently displayed table, refresh it
      if (state.currentDisplayType === 'groupMembers') {
        // Refresh the group members display
        const { members, totalCount } = await fetchAllGroupMembers(
          clearMembersState.selectedGroupId,
          token
        );
        chrome.storage.local.set({ lastGroupMembers: members });
        updateGroupMembersTable(members);
        
        // Update display text
        const displayText = `- ${clearMembersState.selectedGroupName} (${totalCount} members)`;
        document.getElementById('deviceNameDisplay').textContent = displayText;
      }
      
    } catch (error) {
      logMessage(`Error during removal: ${error.message}`);
      showResultsSection({
        total: memberIds.length,
        removed: 0,
        failed: memberIds.length,
        failures: [{
          memberId: 'N/A',
          memberName: 'N/A',
          reason: error.message
        }]
      });
    }
  };

  // Show results section
  const showResultsSection = (results) => {
    document.getElementById('progressSection').style.display = 'none';
    document.getElementById('resultsSection').style.display = 'block';
    
    // Build summary message
    let summaryMsg = '';
    if (results.removed === results.total) {
      summaryMsg = `✓ Successfully removed all ${results.removed} member${results.removed !== 1 ? 's' : ''}.`;
    } else if (results.removed > 0) {
      summaryMsg = `⚠ Partially completed: ${results.removed} removed, ${results.failed} failed.`;
    } else {
      summaryMsg = `✗ Failed to remove members.`;
    }
    
    document.getElementById('resultsSummary').textContent = summaryMsg;
    
    // Show failure details if any (use textContent to prevent XSS)
    const detailsDiv = document.getElementById('resultsDetails');
    detailsDiv.innerHTML = '';
    
    if (results.failures && results.failures.length > 0) {
      results.failures.forEach(failure => {
        const failureItem = document.createElement('div');
        failureItem.className = 'failure-item';
        
        const nameStrong = document.createElement('strong');
        nameStrong.textContent = failure.memberName;
        failureItem.appendChild(nameStrong);
        failureItem.appendChild(document.createElement('br'));
        
        const idSmall = document.createElement('small');
        idSmall.textContent = `ID: ${failure.memberId}`;
        failureItem.appendChild(idSmall);
        failureItem.appendChild(document.createElement('br'));
        
        const reasonSmall = document.createElement('small');
        reasonSmall.textContent = `Reason: ${failure.reason}`;
        failureItem.appendChild(reasonSmall);
        
        detailsDiv.appendChild(failureItem);
      });
    }
  };

  // Handle Clear Group Members button click
  const handleClearGroupMembers = async () => {
    logMessage("clearGroupMembers clicked");
    
    // Try to use cached group info if we're viewing group members
    let groupId, groupName;
    
    if (state.currentDisplayType === 'groupMembers' && state.lastCheckedGroup) {
      // Use cached group info from "Check members"
      groupId = state.lastCheckedGroup.groupId;
      groupName = state.lastCheckedGroup.groupName;
      logMessage(`clearGroupMembers: Using cached group - ID: ${groupId}, Name: ${groupName}`);
    } else {
      // Fall back to requiring selected group from search results
      const selected = document.querySelectorAll("#groupResults input[type=checkbox]:checked");
      if (selected.length !== 1) {
        showResultNotification('Select exactly one group to clear members.', 'error');
        return;
      }
      groupId = selected[0].value;
      groupName = selected[0].dataset.groupName;
      logMessage(`clearGroupMembers: Using selected group - ID: ${groupId}, Name: ${groupName}`);
    }
    
    // Check if group is dynamic
    if (isDynamicGroup(groupId)) {
      showResultNotification(
        'Cannot clear members from dynamic groups. Dynamic membership is managed by Azure AD rules.',
        'error'
      );
      return;
    }

    logMessage(`clearGroupMembers: Processing group - ID: ${groupId}, Name: ${groupName}`);

    try {
      const token = await getToken();
      
      // Fetch current group members
      showProcessingNotification(`Loading members for group "${groupName}"...`);
      
      const { members, totalCount } = await fetchAllGroupMembers(groupId, token);
      
      if (totalCount === 0) {
        showResultNotification(`Group "${groupName}" has no members to clear.`, 'info');
        return;
      }
      
      // Store in state
      clearMembersState.selectedGroupId = groupId;
      clearMembersState.selectedGroupName = groupName;
      clearMembersState.isGroupDynamic = false;
      clearMembersState.allMembers = members;
      
      // Show modal
      document.getElementById('clearMembersGroupName').textContent = 
        `Group: ${groupName}`;
      
      // Update counts
      const selectedMembers = getSelectedMembers();
      updateMemberCounts(selectedMembers.length, totalCount);
      
      showClearMembersModal();
      
      // Hide processing notification
      showResultNotification('', 'clear');
      
    } catch (error) {
      logMessage(`clearGroupMembers: Error - ${error.message}`);
      
      let errorMessage = 'Failed to load group members: ' + error.message;
      
      if (error.message.includes('403') || error.message.includes('Forbidden')) {
        errorMessage = `Access denied. You don't have permission to view or modify members of group "${groupName}".`;
      } else if (error.message.includes('404') || error.message.includes('Not Found')) {
        errorMessage = `Group "${groupName}" was not found or has been deleted.`;
      }
      
      showResultNotification(errorMessage, 'error');
    }
  };

  // Update the button state based on selected group type
  const updateClearMembersButtonState = () => {
    const clearBtn = document.getElementById('clearGroupMembers');
    
    // Check if we have a cached group from "Check members"
    const hasCachedGroup = state.currentDisplayType === 'groupMembers' && state.lastCheckedGroup;
    
    // Check if we have a selected group from search results
    const selected = document.querySelectorAll("#groupResults input[type=checkbox]:checked");
    const hasSelectedGroup = selected.length === 1;
    
    // Button should be enabled if we have either a cached group or a selected group
    if (hasCachedGroup) {
      // Check if the cached group is dynamic
      if (isDynamicGroup(state.lastCheckedGroup.groupId)) {
        clearBtn.classList.add('disabled');
        clearBtn.title = 'Only available for Assigned groups. Dynamic membership cannot be manually cleared.';
      } else {
        clearBtn.classList.remove('disabled');
        clearBtn.title = '';
      }
      return;
    }
    
    if (hasSelectedGroup) {
      const groupId = selected[0].value;
      if (isDynamicGroup(groupId)) {
        clearBtn.classList.add('disabled');
        clearBtn.title = 'Only available for Assigned groups. Dynamic membership cannot be manually cleared.';
      } else {
        clearBtn.classList.remove('disabled');
        clearBtn.title = '';
      }
      return;
    }
    
    // No cached group and no selected group - button stays enabled but will show error when clicked
    clearBtn.classList.remove('disabled');
    clearBtn.title = '';
  };

  // ══════════════════════════════════════════════════════════════
  // End of Clear Group Members Feature
  // ══════════════════════════════════════════════════════════════

  // Handle Checking Group Assignments in Configurations
  const handleCheckGroupAssignments = async () => {
    logMessage("checkGroupAssignments clicked");
    const selected = document.querySelectorAll("#groupResults input[type=checkbox]:checked");
    if (selected.length !== 1) {
      showResultNotification('Select exactly one group to check assignments.', 'error');
      return;
    }

    document.getElementById('profileFilterInput').value = '';
    chrome.storage.local.set({ profileFilterValue: '' });

    clearTableSelection();

    const groupId = selected[0].value;
    const groupName = selected[0].dataset.groupName;

    logMessage(`checkGroupAssignments: Selected group - ID: ${groupId}, Name: ${groupName}`);

    try {
      const token = await getToken();
      logMessage("checkGroupAssignments: Token retrieved successfully");

      showProcessingNotification(`Checking assignments for group "${groupName}"...`);

      // Fetch all configurations using batch requests for optimization
      const allAssignments = [];

      // Batch request 1: Configuration policies and device configurations
      const batch1Body = {
        requests: [
          { id: "deviceConfigurations", method: "GET", url: "/deviceManagement/deviceConfigurations?$select=id,displayName" },
          { id: "configurationPolicies", method: "GET", url: "/deviceManagement/configurationPolicies?$top=500" },
          { id: "groupPolicyConfigurations", method: "GET", url: "/deviceManagement/groupPolicyConfigurations?$top=500" },
          { id: "deviceCompliancePolicies", method: "GET", url: "/deviceManagement/deviceCompliancePolicies?$select=id,displayName" },
          { id: "deviceManagementScripts", method: "GET", url: "/deviceManagement/deviceManagementScripts?$select=id,displayName" }
        ]
      };

      logMessage("checkGroupAssignments: Sending batch request 1...");
      const batch1Response = await fetchJSON("https://graph.microsoft.com/beta/$batch", {
        method: "POST",
        headers: { "Authorization": token, "Content-Type": "application/json" },
        body: JSON.stringify(batch1Body)
      });

      // Process batch 1 responses and collect all configuration IDs
      const configsToCheck = [];
      
      for (const response of batch1Response.responses) {
        if (response.status === 200 && response.body && response.body.value) {
          const configType = getConfigTypeFromBatchId(response.id);
          for (const config of response.body.value) {
            configsToCheck.push({
              id: config.id,
              displayName: config.displayName || config.name || 'Unknown',
              type: configType,
              endpoint: getAssignmentEndpoint(response.id, config.id)
            });
          }
        }
      }

      // Batch request 2: Mobile apps
      logMessage("checkGroupAssignments: Fetching mobile apps...");
      const appsResponse = await fetchJSON("https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?$filter=(microsoft.graph.managedApp/appAvailability eq null or microsoft.graph.managedApp/appAvailability eq 'lineOfBusiness' or isAssigned eq true)&$select=id,displayName&$top=500", {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });

      if (appsResponse.value) {
        for (const app of appsResponse.value) {
          configsToCheck.push({
            id: app.id,
            displayName: app.displayName || 'Unknown App',
            type: 'App',
            endpoint: `/deviceAppManagement/mobileApps/${app.id}/assignments`
          });
        }
      }

      // Batch request 3: Shell scripts (macOS)
      logMessage("checkGroupAssignments: Fetching shell scripts...");
      try {
        const shellScriptsResponse = await fetchJSON("https://graph.microsoft.com/beta/deviceManagement/deviceShellScripts?$select=id,displayName", {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        });

        if (shellScriptsResponse.value) {
          for (const script of shellScriptsResponse.value) {
            configsToCheck.push({
              id: script.id,
              displayName: script.displayName || 'Unknown Script',
              type: 'Shell Script',
              endpoint: `/deviceManagement/deviceShellScripts/${script.id}?$expand=assignments`
            });
          }
        }
      } catch (shellError) {
        logMessage(`checkGroupAssignments: Shell scripts fetch failed (may not be available): ${shellError.message}`);
      }

      // Fetch Security Baselines (Settings Catalog based - new style)
      logMessage("checkGroupAssignments: Fetching security baseline templates...");
      try {
        const securityBaselineTemplatesResponse = await fetchJSON("https://graph.microsoft.com/beta/deviceManagement/configurationPolicyTemplates?$top=500&$filter=(lifecycleState eq 'draft' or lifecycleState eq 'active') and (templateFamily eq 'Baseline')", {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        });

        if (securityBaselineTemplatesResponse.value) {
          // For each baseline template, fetch policies that use it
          for (const template of securityBaselineTemplatesResponse.value) {
            try {
              const baselinePoliciesResponse = await fetchJSON(`https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$select=id,name&$filter=(templateReference/TemplateId eq '${template.id}') and (templateReference/TemplateFamily eq 'Baseline')`, {
                method: "GET",
                headers: { "Authorization": token, "Content-Type": "application/json" }
              });

              if (baselinePoliciesResponse.value) {
                for (const policy of baselinePoliciesResponse.value) {
                  configsToCheck.push({
                    id: policy.id,
                    displayName: policy.name || 'Unknown Security Baseline',
                    type: 'Security Baseline',
                    endpoint: `/deviceManagement/configurationPolicies('${policy.id}')/assignments`
                  });
                }
              }
            } catch (policyError) {
              logMessage(`checkGroupAssignments: Error fetching policies for baseline template ${template.id}: ${policyError.message}`);
            }
          }
        }
      } catch (baselineError) {
        logMessage(`checkGroupAssignments: Security baseline templates fetch failed: ${baselineError.message}`);
      }

      // Fetch Security Baselines (Intent based - legacy style)
      logMessage("checkGroupAssignments: Fetching legacy security baseline intents...");
      const legacyBaselineTemplateIds = [
        '034ccd46-190c-4afc-adf1-ad7cc11262eb', // Windows Security Baseline
        'c04a010a-e7c5-44b1-a814-88df6f053f16', // Windows Security Baseline (older)
        '2209e067-9c8c-462e-9981-5a8c79165dcc', // Defender ATP Baseline
        'a8d6fa0e-1e66-455b-bb51-8ce0dde1559e', // Edge Baseline
        'cef15778-c3b9-4d53-a00a-042929f0aad0'  // Microsoft 365 Apps Baseline
      ];
      
      try {
        const filterConditions = legacyBaselineTemplateIds.map(id => `templateId eq '${id}'`).join(' or ');
        const intentsResponse = await fetchJSON(`https://graph.microsoft.com/beta/deviceManagement/intents?$filter=${encodeURIComponent(filterConditions)}`, {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        });

        if (intentsResponse.value) {
          for (const intent of intentsResponse.value) {
            configsToCheck.push({
              id: intent.id,
              displayName: intent.displayName || 'Unknown Security Baseline',
              type: 'Security Baseline (Legacy)',
              endpoint: `/deviceManagement/intents/${intent.id}/assignments`
            });
          }
        }
      } catch (intentsError) {
        logMessage(`checkGroupAssignments: Legacy security baseline intents fetch failed: ${intentsError.message}`);
      }

      // Fetch Enrollment Restrictions (with assignments expanded)
      logMessage("checkGroupAssignments: Fetching enrollment restrictions...");
      try {
        const enrollmentRestrictionsResponse = await fetchJSON("https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations?$expand=assignments", {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        });

        if (enrollmentRestrictionsResponse.value) {
          for (const config of enrollmentRestrictionsResponse.value) {
            // Determine the type based on deviceEnrollmentConfigurationType or @odata.type
            let configType = 'Enrollment Configuration';
            const odataType = config['@odata.type'] || '';
            const enrollmentType = config.deviceEnrollmentConfigurationType || '';
            
            if (odataType.includes('LimitConfiguration') || enrollmentType === 'limit') {
              configType = 'Device Limit Restriction';
            } else if (odataType.includes('PlatformRestrictions') || enrollmentType === 'platformRestrictions' || enrollmentType === 'singlePlatformRestriction') {
              configType = 'Platform Restriction';
            } else if (odataType.includes('WindowsHelloForBusiness') || enrollmentType === 'windowsHelloForBusiness') {
              configType = 'Windows Hello for Business';
            } else if (odataType.includes('EnrollmentCompletionPage') || enrollmentType === 'windows10EnrollmentCompletionPageConfiguration') {
              configType = 'Enrollment Status Page';
            } else if (odataType.includes('windowsRestore') || enrollmentType === 'windowsRestore') {
              configType = 'Windows Restore';
            }

            // Check if the group is in the assignments (already expanded)
            if (config.assignments && config.assignments.length > 0) {
              for (const assignment of config.assignments) {
                if (!assignment.target) continue;

                const targetType = (assignment.target['@odata.type'] || '').toLowerCase();
                const targetGroupId = assignment.target.groupId;

                if (targetGroupId === groupId) {
                  let intent = 'Included';
                  if (targetType.includes('exclusion')) {
                    intent = 'Excluded';
                  }

                  allAssignments.push({
                    configName: config.displayName || 'Unknown',
                    configType: configType,
                    intent: intent
                  });
                }
              }
            }
          }
        }
      } catch (enrollmentError) {
        logMessage(`checkGroupAssignments: Enrollment restrictions fetch failed: ${enrollmentError.message}`);
      }

      // Fetch Autopilot v2 (Device Preparation) profiles
      logMessage("checkGroupAssignments: Fetching Autopilot v2 (Device Preparation) profiles...");
      try {
        const autopilotV2Response = await fetchJSON("https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$select=id,name&$filter=(technologies has 'enrollment') and (platforms eq 'windows10') and (TemplateReference/templateId eq '80d33118-b7b4-40d8-b15f-81be745e053f_1') and (Templatereference/templateFamily eq 'enrollmentConfiguration')", {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        });

        if (autopilotV2Response.value) {
          for (const profile of autopilotV2Response.value) {
            // Add to configsToCheck for standard assignments
            configsToCheck.push({
              id: profile.id,
              displayName: profile.name || 'Unknown Device Preparation Profile',
              type: 'Autopilot Device Preparation',
              endpoint: `/deviceManagement/configurationPolicies('${profile.id}')/assignments`
            });

            // Also check enrollmentTimeDeviceMembershipTarget for device group assignments
            try {
              const deviceMembershipResponse = await fetchJSON(`https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('${profile.id}')/retrieveEnrollmentTimeDeviceMembershipTarget`, {
                method: "GET",
                headers: { "Authorization": token, "Content-Type": "application/json" }
              });

              if (deviceMembershipResponse.enrollmentTimeDeviceMembershipTargetValidationStatuses) {
                for (const status of deviceMembershipResponse.enrollmentTimeDeviceMembershipTargetValidationStatuses) {
                  if (status.targetId === groupId) {
                    allAssignments.push({
                      configName: profile.name || 'Unknown Device Preparation Profile',
                      configType: 'Autopilot Device Preparation (Device Group)',
                      intent: 'Included'
                    });
                  }
                }
              }
            } catch (deviceMembershipError) {
              logMessage(`checkGroupAssignments: Device membership target fetch failed for profile ${profile.id}: ${deviceMembershipError.message}`);
            }
          }
        }
      } catch (autopilotV2Error) {
        logMessage(`checkGroupAssignments: Autopilot v2 profiles fetch failed: ${autopilotV2Error.message}`);
      }

      // Fetch Autopilot v1 Deployment Profiles (with assignments expanded)
      logMessage("checkGroupAssignments: Fetching Autopilot v1 deployment profiles...");
      try {
        const autopilotV1Response = await fetchJSON("https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles?$expand=assignments&$top=500", {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        });

        if (autopilotV1Response.value) {
          for (const profile of autopilotV1Response.value) {
            // Determine profile type based on @odata.type
            const odataType = profile['@odata.type'] || '';
            let profileType = 'Autopilot Deployment Profile';
            if (odataType.includes('activeDirectory')) {
              profileType = 'Autopilot Hybrid Join Profile';
            } else if (odataType.includes('azureAD')) {
              profileType = 'Autopilot Entra Join Profile';
            }

            // Check if the group is in the assignments (already expanded)
            if (profile.assignments && profile.assignments.length > 0) {
              for (const assignment of profile.assignments) {
                if (!assignment.target) continue;

                const targetType = (assignment.target['@odata.type'] || '').toLowerCase();
                const targetGroupId = assignment.target.groupId;

                if (targetGroupId === groupId) {
                  let intent = 'Included';
                  if (targetType.includes('exclusion')) {
                    intent = 'Excluded';
                  }

                  allAssignments.push({
                    configName: profile.displayName || 'Unknown',
                    configType: profileType,
                    intent: intent
                  });
                }
              }
            }
          }
        }
      } catch (autopilotV1Error) {
        logMessage(`checkGroupAssignments: Autopilot v1 profiles fetch failed: ${autopilotV1Error.message}`);
      }

      logMessage(`checkGroupAssignments: Total configurations to check: ${configsToCheck.length}`);

      // Now fetch assignments for each configuration using batched requests
      const batchSize = 20; // MS Graph batch limit
      let processedCount = 0;

      for (let i = 0; i < configsToCheck.length; i += batchSize) {
        const batch = configsToCheck.slice(i, i + batchSize);
        const batchRequests = batch.map((config, idx) => ({
          id: `${i + idx}`,
          method: "GET",
          url: config.endpoint
        }));

        try {
          const batchResponse = await fetchJSON("https://graph.microsoft.com/beta/$batch", {
            method: "POST",
            headers: { "Authorization": token, "Content-Type": "application/json" },
            body: JSON.stringify({ requests: batchRequests })
          });

          for (const response of batchResponse.responses) {
            const configIndex = parseInt(response.id);
            const config = configsToCheck[configIndex];

            if (response.status === 200) {
              // Handle different response formats
              let assignments = [];
              if (response.body.value) {
                assignments = response.body.value;
              } else if (response.body.assignments) {
                assignments = response.body.assignments;
              }

              // Check if the group is in the assignments
              for (const assignment of assignments) {
                if (!assignment.target) continue;

                const targetType = (assignment.target['@odata.type'] || '').toLowerCase();
                const targetGroupId = assignment.target.groupId;

                if (targetGroupId === groupId) {
                  // Determine intent
                  let intent = 'Included';
                  if (targetType.includes('exclusion')) {
                    intent = 'Excluded';
                  } else if (assignment.intent) {
                    intent = capitalizeFirst(assignment.intent);
                  }

                  allAssignments.push({
                    configName: config.displayName,
                    configType: config.type,
                    intent: intent
                  });
                }
              }
            }
          }
        } catch (batchError) {
          logMessage(`checkGroupAssignments: Batch error at ${i}: ${batchError.message}`);
        }

        processedCount += batch.length;
        if (processedCount % 100 === 0) {
          logMessage(`checkGroupAssignments: Processed ${processedCount}/${configsToCheck.length} configurations`);
        }
      }

      // Clear other data types from storage
      chrome.storage.local.remove(['lastConfigAssignments', 'lastAppAssignments', 'lastComplianceAssignments', 'lastPwshAssignments', 'lastGroupMembers']);
      chrome.storage.local.set({ lastGroupAssignments: allAssignments });

      // Update UI
      document.getElementById('deviceNameDisplay').textContent = `- ${groupName} (${allAssignments.length} assignments)`;

      updateGroupAssignmentsTable(allAssignments);

      if (allAssignments.length === 0) {
        showResultNotification(`Group "${groupName}" is not assigned to any configurations.`, 'info');
      } else {
        showResultNotification(`Found ${allAssignments.length} configuration assignments for group "${groupName}".`, 'success');
      }

    } catch (error) {
      logMessage(`checkGroupAssignments: Error - ${error.message}`);
      showResultNotification('Failed to check group assignments: ' + error.message, 'error');
    }
  };

  // Helper function to get config type from batch ID
  const getConfigTypeFromBatchId = (batchId) => {
    const typeMap = {
      'deviceConfigurations': 'Device Configuration',
      'configurationPolicies': 'Settings Catalog',
      'groupPolicyConfigurations': 'Administrative Templates',
      'deviceCompliancePolicies': 'Compliance Policy',
      'deviceManagementScripts': 'PowerShell Script'
    };
    return typeMap[batchId] || 'Configuration';
  };

  // Helper function to get assignment endpoint
  const getAssignmentEndpoint = (batchId, configId) => {
    const endpointMap = {
      'deviceConfigurations': `/deviceManagement/deviceConfigurations/${configId}/assignments`,
      'configurationPolicies': `/deviceManagement/configurationPolicies('${configId}')/assignments`,
      'groupPolicyConfigurations': `/deviceManagement/groupPolicyConfigurations/${configId}/assignments`,
      'deviceCompliancePolicies': `/deviceManagement/deviceCompliancePolicies/${configId}/assignments`,
      'deviceManagementScripts': `/deviceManagement/deviceManagementScripts/${configId}/assignments`
    };
    return endpointMap[batchId] || `/deviceManagement/${batchId}/${configId}/assignments`;
  };

  // Helper function to capitalize first letter
  const capitalizeFirst = (str) => {
    if (!str) return '';
    return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
  };

  // Handle Checking Configuration Assignments
  const handleCheckGroups = async () => {
    logMessage("checkGroups clicked");
    showProcessingNotification('Fetching configuration assignments...');
    document.getElementById('profileFilterInput').value = '';
    chrome.storage.local.set({ profileFilterValue: '' });

    chrome.storage.local.remove(['lastAppAssignments', 'lastComplianceAssignments', 'lastPwshAssignments', 'lastGroupMembers']);

    // Clear table selection before loading new assignments
    clearTableSelection();

    // Clear dynamic groups tracking before fetching new assignments
    clearDynamicGroups();

    try {
      const { mdmDeviceId } = await verifyMdmUrl();
      const token = await getToken();
      const deviceData = await fetchJSON(`https://graph.microsoft.com/beta/deviceManagement/manageddevices('${mdmDeviceId}')`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });

      // Update device name display
      updateDeviceNameDisplay(deviceData);

      const azureADDeviceId = deviceData.azureADDeviceId;
      const userPrincipalName = deviceData.userPrincipalName;
      if (!azureADDeviceId) throw new Error("Could not find Azure AD Device ID for this device.");
      const filter = encodeURIComponent(`deviceId eq '${azureADDeviceId}'`);
      const deviceObjData = await fetchJSON(`https://graph.microsoft.com/beta/devices?$filter=${filter}`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });
      if (!deviceObjData.value || deviceObjData.value.length === 0) throw new Error("Could not find the device in Azure AD.");
      const deviceObjectId = deviceObjData.value[0].id;
      let userPromise;
      if (userPrincipalName && userPrincipalName !== 'Unknown user') {
        userPromise = fetchJSON(`https://graph.microsoft.com/beta/users?$filter=userPrincipalName eq '${encodeURIComponent(userPrincipalName)}'`, {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        }).then(userData => (userData.value && userData.value.length > 0) ? userData.value[0].id : null);
      } else {
        userPromise = Promise.resolve(null);
      }
      const userObjectId = await userPromise;
      const groupMaps = await getAllGroupsMap(deviceObjectId, userObjectId, token);
      const reportBody = JSON.stringify({
        top: "500",
        skip: "0",
        select: ["PolicyId", "PolicyName", "PolicyType", "UPN"],
        filter: `((PolicyBaseTypeName eq 'Microsoft.Management.Services.Api.DeviceConfiguration') or (PolicyBaseTypeName eq 'DeviceManagementConfigurationPolicy') or (PolicyBaseTypeName eq 'DeviceConfigurationAdmxPolicy') or (PolicyBaseTypeName eq 'Microsoft.Management.Services.Api.DeviceManagementIntent')) and (IntuneDeviceId eq '${mdmDeviceId}')`
      });
      const reportData = await fetchJSON("https://graph.microsoft.com/beta/deviceManagement/reports/getConfigurationPoliciesReportForDevice", {
        method: "POST",
        headers: { "Authorization": token, "Content-Type": "application/json" },
        body: reportBody
      });
      const policies = reportData.Values || [];
      if (policies.length === 0) {
        showResultNotification('No policies found.', 'info');
        clearTableAndPagination();
        return;
      }
      const assignmentsPromises = policies.map(async (row) => {
        const policy = {
          PolicyId: row[0],
          PolicyName: row[1],
          PolicyType: row[2],
          UPN: row[3]
        };
        const specialPolicyIds = ["26", "20", "33", "55", "118", "75", "72", "25", "31", "107", "99999"];
        const endpoint = specialPolicyIds.includes(policy.PolicyType.toString())
          ? `https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/${policy.PolicyId}/assignments`
          : `https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('${policy.PolicyId}')/assignments`;
        try {
          const assignData = await fetchJSON(endpoint, {
            method: "GET",
            headers: { "Authorization": token, "Content-Type": "application/json" }
          });
          let targetObjs = [];
          const assignments = Array.isArray(assignData.value) ? assignData.value : (assignData.value ? [assignData.value] : []);
          // Process assignments sequentially to ensure proper group resolution
          for (const asg of assignments) {
            if (!asg.target) continue;
            const typeRaw = (asg.target['@odata.type'] || "").toLowerCase().trim();
            if (typeRaw.includes("groupassignmenttarget")) {
              const groupId = asg.target.groupId;
              let groupName, membershipType;
              
              if (groupMaps.allGroups.has(groupId)) {
                // Group found in initial mapping
                groupName = groupMaps.allGroups.get(groupId);
                
                // Determine membership type based on direct vs transitive membership
                const isDirectMember = groupMaps.directGroups.has(groupId);
                const isTransitiveMember = groupMaps.transitiveGroups.has(groupId);
                
                if (!isDirectMember && isTransitiveMember) {
                  membershipType = 'Transitive';
                  logMessage(`checkGroups: Group ${groupName} (${groupId}) - Transitive membership detected`);
                } else if (isDirectMember) {
                  membershipType = 'Direct';
                  logMessage(`checkGroups: Group ${groupName} (${groupId}) - Direct membership detected`);
                } else {
                  membershipType = 'Direct'; // Default fallback
                }
              } else {
                // Group not found in initial mapping - try to resolve it
                logMessage(`checkGroups: Group ${groupId} not found in groupMaps, attempting to resolve...`);
                const resolvedGroup = await resolveGroupInfo(groupId, deviceObjectId, userObjectId, token, groupMaps);
                
                if (resolvedGroup) {
                  groupName = resolvedGroup.groupName;
                  membershipType = resolvedGroup.membershipType;
                  logMessage(`checkGroups: Successfully resolved missing group ${groupName} (${groupId})`);
                } else {
                  // Group exists in assignment but device/user is not a member - correctly filter out
                  logMessage(`checkGroups: Group ${groupId} assignment filtered out (device/user not a member)`);
                  continue;
                }
              }
              
              targetObjs.push({
                groupId: groupId,
                groupName: groupName,
                membershipType: membershipType,
                targetType: typeRaw.includes('user') ? 'User' : 'Device',
                intent: "Included"
              });
            } else if (typeRaw.includes("alldevicesassignmenttarget")) {
              targetObjs.push({
                groupName: "All Devices",
                membershipType: "Virtual",
                targetType: "Device",
                intent: "Included"
              });
            } else if (typeRaw.includes("allusersassignmenttarget")) {
              if (policy.UPN && policy.UPN !== 'Not Available') {
                targetObjs.push({
                  groupName: "All Users",
                  membershipType: "Virtual",
                  targetType: "User",
                  intent: "Included"
                });
              }
            }
          }
          return { policyName: policy.PolicyName, targets: targetObjs };
        } catch (err) {
          logMessage(`checkGroups: Error processing policy ${policy.PolicyName} - ${err.message}`);
          return { policyName: policy.PolicyName, targets: [] };
        }
      });
      const finalResults = (await Promise.all(assignmentsPromises)).filter(result => result.targets && result.targets.length > 0);
      chrome.storage.local.set({ lastConfigAssignments: finalResults });
      updateConfigTable(finalResults);
      logMessage(`checkGroups: Found ${finalResults.length} policies with valid group assignments`);
      showResultNotification('Configuration assignments loaded successfully', 'success');
    } catch (error) {
      logMessage(`checkGroups: Error - ${error.message}`);
      showResultNotification('Failed to load configuration assignments: ' + error.message, 'error');
    }
  };
  // Handle Checking Compliance Policies
  const handleCheckCompliance = async () => {
    logMessage("checkCompliance clicked");
    showProcessingNotification('Fetching compliance policies...');
    document.getElementById('profileFilterInput').value = '';
    chrome.storage.local.set({ profileFilterValue: '' });

    chrome.storage.local.remove(['lastConfigAssignments','lastAppAssignments','lastPwshAssignments','lastGroupMembers']);

    // Clear table selection before loading new assignments
    clearTableSelection();
    let token, mdmDeviceId;
    try {
      ({ mdmDeviceId } = await verifyMdmUrl());
      token = await getToken();
      const deviceData = await fetchJSON(`https://graph.microsoft.com/beta/deviceManagement/manageddevices('${mdmDeviceId}')`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });

      // Update device name display
      updateDeviceNameDisplay(deviceData);

      const azureADDeviceId = deviceData.azureADDeviceId;
      const userPrincipalName = deviceData.userPrincipalName;
      if (!azureADDeviceId) throw new Error("Could not find Azure AD Device ID for this device.");
      const filter = encodeURIComponent(`deviceId eq '${azureADDeviceId}'`);
      const deviceObjData = await fetchJSON(`https://graph.microsoft.com/beta/devices?$filter=${filter}`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });
      if (!deviceObjData.value || deviceObjData.value.length === 0) throw new Error("Could not find the device in Azure AD.");
      const deviceObjectId = deviceObjData.value[0].id;
      let userPromise;
      if (userPrincipalName && userPrincipalName !== 'Unknown user') {
        userPromise = fetchJSON(`https://graph.microsoft.com/beta/users?$filter=userPrincipalName eq '${encodeURIComponent(userPrincipalName)}'`, {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        }).then(userData => (userData.value && userData.value.length > 0) ? userData.value[0].id : null);
      } else {
        userPromise = Promise.resolve(null);
      }
      const userObjectId = await userPromise;
      const groupMaps = await getAllGroupsMap(deviceObjectId, userObjectId, token);
      
      // Try a comprehensive approach: fetch all compliance policies from tenant and match them to device/user groups
      logMessage("checkCompliance: Fetching all compliance policies from tenant to cross-reference with device report");
      let tenantCompliancePolicies = [];
      try {
        const allPoliciesData = await fetchJSON("https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies?$expand=assignments", {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        });
        tenantCompliancePolicies = allPoliciesData.value || [];
        logMessage(`checkCompliance: Found ${tenantCompliancePolicies.length} compliance policies in tenant`);
      } catch (tenantErr) {
        logMessage(`checkCompliance: Failed to fetch tenant compliance policies: ${tenantErr.message}`);
      }
      
      const reportBody = JSON.stringify({
        filter: `(DeviceId eq '${mdmDeviceId}') and ((PolicyPlatformType eq '4') or (PolicyPlatformType eq '5') or (PolicyPlatformType eq '6') or (PolicyPlatformType eq '8') or (PolicyPlatformType eq '100'))`,
        orderBy: ["PolicyName asc"]
      });
      const reportData = await fetchJSON("https://graph.microsoft.com/beta/deviceManagement/reports/getDevicePoliciesComplianceReport", {
        method: "POST",
        headers: { "Authorization": token, "Content-Type": "application/json" },
        body: reportBody
      });
      if (!reportData || !reportData.Schema || !reportData.Values) {
        logMessage("checkCompliance: Invalid response format");
        showResultNotification('Invalid response format from compliance report API.', 'error');
        return;
      }
      const schemaMap = {};
      reportData.Schema.forEach((col, idx) => { schemaMap[col.Column] = idx; });
      const requiredColumns = ['PolicyId', 'PolicyName', 'PolicyStatus_loc'];
      const missingColumns = requiredColumns.filter(col => schemaMap[col] === undefined);
      if (missingColumns.length > 0) {
        logMessage(`checkCompliance: Missing required columns: ${missingColumns.join(', ')}`);
        showResultNotification(`API response missing required columns: ${missingColumns.join(', ')}`, 'error');
        return;
      }
      const policies = reportData.Values || [];
      if (policies.length === 0) {
        logMessage("checkCompliance: No compliance policies found");
        showResultNotification('No compliance policies found.', 'info');
        clearTableAndPagination();
        return;
      }
      const formattedPolicies = policies.map(policy => ({
        policyId: policy[schemaMap.PolicyId] || 'Unknown',
        policyName: policy[schemaMap.PolicyName] || 'Unknown Policy',
        complianceState: policy[schemaMap.PolicyStatus_loc] || 'Unknown'
      }));
      
      logMessage(`checkCompliance: Found ${formattedPolicies.length} policies from device report`);
      formattedPolicies.forEach((policy, index) => {
        logMessage(`checkCompliance: Policy ${index + 1}: "${policy.policyName}" (ID: ${policy.policyId}) - State: ${policy.complianceState}`);
      });
      
      const assignmentPromises = formattedPolicies.map(async (policy) => {
        if (!policy.policyId || policy.policyId === 'Unknown') {
          logMessage(`checkCompliance: Skipping policy "${policy.policyName}" - no valid policy ID`);
          return { ...policy, assignments: [] };
        }
        try {
          logMessage(`checkCompliance: Fetching assignments for policy "${policy.policyName}" (${policy.policyId})`);
          const policyData = await fetchJSON(`https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies/${policy.policyId}?$expand=assignments`, {
            method: "GET",
            headers: { "Authorization": token, "Content-Type": "application/json" }
          });
          
          const assignmentCount = policyData.assignments ? policyData.assignments.length : 0;
          logMessage(`checkCompliance: Policy "${policy.policyName}" has ${assignmentCount} assignments`);
          
          return { ...policy, assignments: policyData.assignments || [] };
        } catch (err) {
          logMessage(`checkCompliance: Error getting assignments for policy ${policy.policyId}: ${err.message}`);
          
          // If it's a 404 error, try alternative endpoints for different compliance policy types
          if (err.message.includes('404') || err.message.includes('Not Found')) {
            logMessage(`checkCompliance: Policy "${policy.policyName}" (${policy.policyId}) not found in deviceCompliancePolicies endpoint - trying alternative approaches`);
            
            // Try different compliance policy endpoints
            const alternativeEndpoints = [
              `https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies('${policy.policyId}')?$expand=assignments`,
              `https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations('${policy.policyId}')?$expand=assignments`,
              `https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('${policy.policyId}')?$expand=assignments`
            ];
            
            for (const endpoint of alternativeEndpoints) {
              try {
                logMessage(`checkCompliance: Trying alternative endpoint for policy "${policy.policyName}": ${endpoint.substring(0, 100)}...`);
                const altPolicyData = await fetchJSON(endpoint, {
                  method: "GET",
                  headers: { "Authorization": token, "Content-Type": "application/json" }
                });
                
                if (altPolicyData && altPolicyData.assignments) {
                  const assignmentCount = altPolicyData.assignments.length;
                  logMessage(`checkCompliance: Found ${assignmentCount} assignments for policy "${policy.policyName}" using alternative endpoint`);
                  return { ...policy, assignments: altPolicyData.assignments || [] };
                }
              } catch (altErr) {
                logMessage(`checkCompliance: Alternative endpoint failed for policy "${policy.policyName}": ${altErr.message}`);
                continue;
              }
            }
            
            logMessage(`checkCompliance: All endpoints failed for policy "${policy.policyName}" - might be a built-in/system policy`);
          }
          
          return { ...policy, assignments: [] };
        }
      });
      const policiesWithAssignments = await Promise.all(assignmentPromises);
      if (!policiesWithAssignments || policiesWithAssignments.length === 0) {
        clearTableAndPagination();
        showResultNotification('No compliance policies found.', 'info');
        return;
      }
      
      // Process assignments with improved group resolution
      const tableData = await Promise.all(policiesWithAssignments.map(async (policy) => {
        logMessage(`checkCompliance: Processing assignments for policy "${policy.policyName}" with ${policy.assignments ? policy.assignments.length : 0} assignments`);
        let targets = [];
        if (policy.assignments && policy.assignments.length > 0) {
          // Process assignments sequentially to avoid overwhelming the API
          for (const asg of policy.assignments) {
            if (!asg.target) {
              logMessage(`checkCompliance: Policy "${policy.policyName}" - skipping assignment with no target`);
              continue;
            }
            const targetType = (asg.target['@odata.type'] || '').toLowerCase();
            logMessage(`checkCompliance: Policy "${policy.policyName}" - processing assignment with target type: ${targetType}`);
            
            const isExclusion = targetType.includes('exclusion');
            if (isExclusion) {
              logMessage(`checkCompliance: Policy "${policy.policyName}" - skipping exclusion assignment`);
              continue; // Skip exclusions
            }

            if (targetType.includes('groupassignmenttarget')) {
              const groupId = asg.target.groupId;
              logMessage(`checkCompliance: Policy "${policy.policyName}" - processing group assignment for group ${groupId}`);
              let groupName, membershipType;
              
              if (groupMaps.allGroups.has(groupId)) {
                // Group found in initial mapping
                groupName = groupMaps.allGroups.get(groupId);
                logMessage(`checkCompliance: Policy "${policy.policyName}" - group ${groupName} (${groupId}) found in groupMaps`);
                
                // Determine membership type based on direct vs transitive membership
                const isDirectMember = groupMaps.directGroups.has(groupId);
                const isTransitiveMember = groupMaps.transitiveGroups.has(groupId);
                
                if (!isDirectMember && isTransitiveMember) {
                  membershipType = 'Transitive';
                  logMessage(`checkCompliance: Group ${groupName} (${groupId}) - Transitive membership detected`);
                } else if (isDirectMember) {
                  membershipType = 'Direct';
                  logMessage(`checkCompliance: Group ${groupName} (${groupId}) - Direct membership detected`);
                } else {
                  membershipType = 'Direct'; // Default fallback
                  logMessage(`checkCompliance: Group ${groupName} (${groupId}) - using Direct as fallback`);
                }
              } else {
                // Group not found in initial mapping - try to resolve it
                logMessage(`checkCompliance: Group ${groupId} not found in groupMaps, attempting to resolve...`);
                const resolvedGroup = await resolveGroupInfo(groupId, deviceObjectId, userObjectId, token, groupMaps);
                
                if (resolvedGroup) {
                  groupName = resolvedGroup.groupName;
                  membershipType = resolvedGroup.membershipType;
                  logMessage(`checkCompliance: Successfully resolved missing group ${groupName} (${groupId})`);
                } else {
                  // Group exists in assignment but device/user is not a member - correctly filter out
                  logMessage(`checkCompliance: Group ${groupId} assignment filtered out (device/user not a member)`);
                  continue;
                }
              }

              const targetInfo = {
                groupName,
                membershipType: membershipType,
                targetType: targetType.includes('user') ? 'User' : 'Device'
              };
              targets.push(targetInfo);
              logMessage(`checkCompliance: Policy "${policy.policyName}" - added target: ${JSON.stringify(targetInfo)}`);
            } else if (targetType.includes('alldevicesassignmenttarget')) {
              const targetInfo = {
                groupName: 'All Devices',
                membershipType: 'Virtual',
                targetType: 'Device'
              };
              targets.push(targetInfo);
              logMessage(`checkCompliance: Policy "${policy.policyName}" - added All Devices target: ${JSON.stringify(targetInfo)}`);
            } else if (targetType.includes('allusersassignmenttarget') || targetType.includes('alllicensedusersassignmenttarget')) {
              const targetInfo = {
                groupName: 'All Users',
                membershipType: 'Virtual',
                targetType: 'User'
              };
              targets.push(targetInfo);
              logMessage(`checkCompliance: Policy "${policy.policyName}" - added All Users target: ${JSON.stringify(targetInfo)}`);
            } else {
              logMessage(`checkCompliance: Policy "${policy.policyName}" - unknown target type: ${targetType}`);
            }
          }
        }
        
        logMessage(`checkCompliance: Policy "${policy.policyName}" - final targets count: ${targets.length}`);
        if (targets.length === 0) {
          targets.push({ groupName: 'No Assignments', membershipType: '-', targetType: '-' });
          logMessage(`checkCompliance: Policy "${policy.policyName}" - added No Assignments fallback`);
        }
        return { policyName: policy.policyName, complianceState: policy.complianceState, targets };
      }));
      
      // Cross-reference with tenant compliance policies to find any missing assignments
      logMessage("checkCompliance: Cross-referencing with tenant policies for additional assignments");
      const additionalPolicies = [];
      
      for (const tenantPolicy of tenantCompliancePolicies) {
        // Check if this tenant policy is already in our results
        const alreadyIncluded = tableData.some(reportPolicy => 
          reportPolicy.policyName === tenantPolicy.displayName || 
          reportPolicy.policyName === tenantPolicy.name ||
          (reportPolicy.targets && reportPolicy.targets.length > 0 && reportPolicy.targets[0].groupName !== 'No Assignments')
        );
        
        if (alreadyIncluded) {
          continue;
        }
        
        // Check if this tenant policy is assigned to any groups the device/user is a member of
        if (tenantPolicy.assignments && tenantPolicy.assignments.length > 0) {
          const targets = [];
          
          for (const asg of tenantPolicy.assignments) {
            if (!asg.target) continue;
            const targetType = (asg.target['@odata.type'] || '').toLowerCase();
            const isExclusion = targetType.includes('exclusion');
            if (isExclusion) continue;

            if (targetType.includes('groupassignmenttarget')) {
              const groupId = asg.target.groupId;
              let groupName, membershipType;
              
              if (groupMaps.allGroups.has(groupId)) {
                groupName = groupMaps.allGroups.get(groupId);
                const isDirectMember = groupMaps.directGroups.has(groupId);
                const isTransitiveMember = groupMaps.transitiveGroups.has(groupId);
                
                if (!isDirectMember && isTransitiveMember) {
                  membershipType = 'Transitive';
                } else if (isDirectMember) {
                  membershipType = 'Direct';
                } else {
                  membershipType = 'Direct';
                }
              } else {
                const resolvedGroup = await resolveGroupInfo(groupId, deviceObjectId, userObjectId, token, groupMaps);
                if (resolvedGroup) {
                  groupName = resolvedGroup.groupName;
                  membershipType = resolvedGroup.membershipType;
                } else {
                  continue;
                }
              }

              targets.push({
                groupName,
                membershipType: membershipType,
                targetType: targetType.includes('user') ? 'User' : 'Device'
              });
            } else if (targetType.includes('alldevicesassignmenttarget')) {
              targets.push({
                groupName: 'All Devices',
                membershipType: 'Virtual',
                targetType: 'Device'
              });
            } else if (targetType.includes('allusersassignmenttarget')) {
              targets.push({
                groupName: 'All Users',
                membershipType: 'Virtual',
                targetType: 'User'
              });
            }
          }
          
          if (targets.length > 0) {
            logMessage(`checkCompliance: Found additional tenant policy "${tenantPolicy.displayName}" with ${targets.length} applicable assignments`);
            additionalPolicies.push({
              policyName: tenantPolicy.displayName || tenantPolicy.name || 'Unknown Policy',
              complianceState: 'Not in Device Report',
              targets
            });
          }
        }
      }
      
      // Merge additional policies with the original results
      const finalTableData = [...tableData, ...additionalPolicies];
      logMessage(`checkCompliance: Final results - ${tableData.length} from device report + ${additionalPolicies.length} from tenant = ${finalTableData.length} total policies`);
      
      chrome.storage.local.set({ lastComplianceAssignments: finalTableData });
      updateComplianceTable(finalTableData);
      logMessage(`checkCompliance: Found ${finalTableData.length} compliance policies with assignments`);
      showResultNotification('Compliance policies loaded successfully', 'success');
    } catch (error) {
      logMessage(`checkCompliance: Error - ${error.message}`);
      showResultNotification('Failed to load compliance policies: ' + error.message, 'error');
    }
  };

  // Handle Downloading Script
  const handleDownloadScript = async () => {
    logMessage("downloadScript clicked");
    chrome.tabs.query({ active: true, currentWindow: true }, async (tabs) => {
      if (!tabs || !tabs[0]) {
        logMessage("downloadScript: No active tab found");
        showResultNotification('No active tab found.', 'error');
        return;
      }
      const url = tabs[0].url;
      const policyMatch = url.match(/policyId\/([\w-]+)/);
      if (!policyMatch || !policyMatch[1]) {
        logMessage("downloadScript: No policyId found in URL");
        showResultNotification('Could not find policy ID in the current URL.', 'error');
        return;
      }
      const policyId = policyMatch[1];
      logMessage(`downloadScript: Found policyId: ${policyId}`);
      try {
        const token = await getToken();
        const scriptData = await fetchJSON(`https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts/${policyId}`, {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        });
        if (!scriptData.scriptContent) {
          throw new Error("No script content found in response");
        }
        const decodedScript = atob(scriptData.scriptContent);
        const blob = new Blob([decodedScript], { type: 'text/plain' });
        const blobUrl = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = blobUrl;
        a.download = scriptData.fileName || 'IntuneScript.ps1';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(blobUrl);
        document.body.removeChild(a);
        showResultNotification('Script downloaded successfully', 'success');
      } catch (error) {
        logMessage(`downloadScript: Error - ${error.message}`);
        showResultNotification('Failed to download script: ' + error.message, 'error');
      }
    });
  };
  // Handle App Assignments
  const handleAppsAssignment = async () => {
    logMessage("appsAssignment clicked");
    showProcessingNotification('Fetching app assignments...');
    document.getElementById('profileFilterInput').value = '';
    chrome.storage.local.set({ profileFilterValue: '' });

    chrome.storage.local.remove(['lastConfigAssignments','lastComplianceAssignments','lastPwshAssignments','lastGroupMembers']);

    // Clear table selection before loading new assignments
    clearTableSelection();
    let token;
    try {
      const { mdmDeviceId } = await verifyMdmUrl();
      token = await getToken();
      const deviceData = await fetchJSON(`https://graph.microsoft.com/beta/deviceManagement/manageddevices('${mdmDeviceId}')`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });

      // Update device name display
      updateDeviceNameDisplay(deviceData);

      const azureADDeviceId = deviceData.azureADDeviceId;
      const userPrincipalName = deviceData.userPrincipalName;
      if (!azureADDeviceId) throw new Error("Could not find Azure AD Device ID for this device.");
      const filter = encodeURIComponent(`deviceId eq '${azureADDeviceId}'`);
      const deviceObjData = await fetchJSON(`https://graph.microsoft.com/beta/devices?$filter=${filter}`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });
      if (!deviceObjData.value || deviceObjData.value.length === 0) throw new Error("Could not find the device in Azure AD.");
      const deviceObjectId = deviceObjData.value[0].id;
      let userObjectId = null;
      if (userPrincipalName && userPrincipalName !== 'Unknown user') {
        const userData = await fetchJSON(`https://graph.microsoft.com/beta/users?$filter=userPrincipalName eq '${encodeURIComponent(userPrincipalName)}'`, {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        });
        if (userData.value && userData.value.length > 0) {
          userObjectId = userData.value[0].id;
        }
      }
      const groupMaps = await getAllGroupsMap(deviceObjectId, userObjectId, token);
      const appRequests = [];
      appRequests.push(
        fetchJSON(`https://graph.microsoft.com/beta/users('00000000-0000-0000-0000-000000000000')/mobileAppIntentAndStates('${mdmDeviceId}')`, {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        }).then(data => ({ type: 'Device', apps: data.mobileAppList || [] }))
      );
      if (userObjectId) {
        appRequests.push(
          fetchJSON(`https://graph.microsoft.com/beta/users('${userObjectId}')/mobileAppIntentAndStates('${mdmDeviceId}')`, {
            method: "GET",
            headers: { "Authorization": token, "Content-Type": "application/json" }
          }).then(data => ({ type: 'User', apps: data.mobileAppList || [] }))
        );
      }
      const appResults = await Promise.all(appRequests);
      const allApps = [];
      let totalApps = 0;
      appResults.forEach(result => {
        totalApps += result.apps.length;
        result.apps.forEach(app => {
          allApps.push({
            applicationId: app.applicationId,
            displayName: app.displayName,
            mobileAppIntent: app.mobileAppIntent,
            displayVersion: app.displayVersion || 'N/A',
            installState: app.installState,
            targetType: result.type
          });
        });
      });
      logMessage(`appsAssignment: Found ${totalApps} total apps`);
      const assignmentPromises = allApps.map(async (app) => {
        const assignmentData = await fetchJSON(`https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/${app.applicationId}/assignments`, {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        }).catch(() => ({ value: [] }));
        const assignments = assignmentData.value || [];
        const validTargets = [];
        
        // Process assignments sequentially to avoid overwhelming the API
        for (const assignment of assignments) {
          if (!assignment.target) continue;
          const intentInfo = assignment.intent || app.mobileAppIntent;
          const typeRaw = (assignment.target['@odata.type'] || '').toLowerCase();

          // Skip exclusions
          if (typeRaw.includes('exclusion')) continue;

          if (typeRaw.includes('alldevicesassignmenttarget')) {
            validTargets.push({
              groupName: 'All Devices',
              membershipType: 'Virtual',
              targetType: 'Device',
              intent: intentInfo
            });
          } else if (typeRaw.includes('allusersassignmenttarget') || typeRaw.includes('alllicensedusersassignmenttarget')) {
            validTargets.push({
              groupName: 'All Users',
              membershipType: 'Virtual',
              targetType: 'User',
              intent: intentInfo
            });
          } else if (typeRaw.includes('groupassignmenttarget')) {
            const groupId = assignment.target.groupId;
            let groupName, membershipType;
            
            if (groupMaps.allGroups.has(groupId)) {
              // Group found in initial mapping
              groupName = groupMaps.allGroups.get(groupId);
              
              // Determine membership type based on direct vs transitive membership
              const isDirectMember = groupMaps.directGroups.has(groupId);
              const isTransitiveMember = groupMaps.transitiveGroups.has(groupId);
              
              if (!isDirectMember && isTransitiveMember) {
                membershipType = 'Transitive';
                logMessage(`appsAssignment: Group ${groupName} (${groupId}) - Transitive membership detected`);
              } else if (isDirectMember) {
                membershipType = 'Direct';
                logMessage(`appsAssignment: Group ${groupName} (${groupId}) - Direct membership detected`);
              } else {
                membershipType = 'Direct'; // Default fallback
              }
            } else {
              // Group not found in initial mapping - try to resolve it
              logMessage(`appsAssignment: Group ${groupId} not found in groupMaps, attempting to resolve...`);
              const resolvedGroup = await resolveGroupInfo(groupId, deviceObjectId, userObjectId, token, groupMaps);
              
              if (resolvedGroup) {
                groupName = resolvedGroup.groupName;
                membershipType = resolvedGroup.membershipType;
                logMessage(`appsAssignment: Successfully resolved missing group ${groupName} (${groupId})`);
              } else {
                // Group exists in assignment but device/user is not a member - correctly filter out
                logMessage(`appsAssignment: Group ${groupId} assignment filtered out (device/user not a member)`);
                continue;
              }
            }

            validTargets.push({
              groupId,
              groupName,
              membershipType: membershipType,
              targetType: typeRaw.includes('user') ? 'User' : 'Device',
              intent: intentInfo
            });
          } else {
            validTargets.push({
              groupName: typeRaw,
              membershipType: '-',
              targetType: '-',
              intent: intentInfo
            });
          }
        }
        
        if (validTargets.length === 0) {
          validTargets.push({
            groupName: 'No Assignments',
            membershipType: '-',
            targetType: app.targetType,
            intent: app.mobileAppIntent || '-'
          });
        }
        return {
          appName: app.displayName,
          appVersion: app.displayVersion,
          installState: app.installState,
          targets: validTargets
        };
      });
      const appAssignments = await Promise.all(assignmentPromises);
      chrome.storage.local.set({ lastAppAssignments: appAssignments });
      updateAppTable(appAssignments);
      logMessage(`appsAssignment: Found ${appAssignments.length} apps total`);
      showResultNotification('App assignments loaded successfully', 'success');
    } catch (error) {
      logMessage(`appsAssignment: Error - ${error.message}`);
      showResultNotification('Failed to load app assignments: ' + error.message, 'error');
    }
  };
  // Handle PowerShell Profiles (scripts)
  const handlePwshProfiles = async () => {
    logMessage("pwshProfiles clicked");
    showProcessingNotification('Fetching PowerShell profiles...');
    document.getElementById('profileFilterInput').value = '';
    chrome.storage.local.set({ profileFilterValue: '' });

    chrome.storage.local.remove(['lastConfigAssignments','lastAppAssignments','lastComplianceAssignments','lastGroupMembers']);

    // Clear table selection before loading new assignments
    clearTableSelection();
    let token;
    try {
      const { mdmDeviceId } = await verifyMdmUrl();
      token = await getToken();
      const deviceData = await fetchJSON(`https://graph.microsoft.com/beta/deviceManagement/manageddevices('${mdmDeviceId}')`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });

      // Update device name display
      updateDeviceNameDisplay(deviceData);

      const azureADDeviceId = deviceData.azureADDeviceId;
      const userPrincipalName = deviceData.userPrincipalName;
      if (!azureADDeviceId) throw new Error("Could not find Azure AD Device ID for this device.");
      const filter = encodeURIComponent(`deviceId eq '${azureADDeviceId}'`);
      const deviceObjData = await fetchJSON(`https://graph.microsoft.com/beta/devices?$filter=${filter}`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });
      if (!deviceObjData.value || deviceObjData.value.length === 0) throw new Error("Could not find the device in Azure AD.");
      const deviceObjectId = deviceObjData.value[0].id;
      let userObjectId = null;
      if (userPrincipalName && userPrincipalName !== 'Unknown user') {
        const userData = await fetchJSON(`https://graph.microsoft.com/beta/users?$filter=userPrincipalName eq '${encodeURIComponent(userPrincipalName)}'`, {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        });
        if (userData.value && userData.value.length > 0) userObjectId = userData.value[0].id;
      }
      const groupMaps = await getAllGroupsMap(deviceObjectId, userObjectId, token);
      const scriptsData = await fetchJSON("https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts?$expand=assignments", {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });
      if (!scriptsData.value || scriptsData.value.length === 0) {
        showResultNotification('No PowerShell profiles found.', 'info');
        clearTableAndPagination();
        return;
      }
      const matchedScripts = [];
      let matchCount = 0;
      
      // Process scripts sequentially for better error handling and group resolution
      for (const script of scriptsData.value) {
        if (!script.assignments || script.assignments.length === 0) {
          logMessage(`pwshProfiles: Script "${script.displayName}" has no assignments - skipping`);
          continue;
        }
        
        for (const asg of script.assignments) {
          if (!asg.target) {
            logMessage(`pwshProfiles: Script "${script.displayName}" has assignment without target - skipping`);
            continue;
          }
          let targetName = '';
          let targetTypeInfo = '';
          let isMatch = false;
          
          if (asg.target.groupId) {
            const groupId = asg.target.groupId;
            
            if (groupMaps.allGroups.has(groupId)) {
              targetName = groupMaps.allGroups.get(groupId);
              // Determine if it's a device or user group based on current target mode
              targetTypeInfo = state.targetMode === 'device' ? 'Device' : 'User';
              isMatch = true;
              matchCount++;
              logMessage(`pwshProfiles: MATCH - Group ${targetName}`);
            } else {
              // Group not found in initial mapping - try to resolve it
              logMessage(`pwshProfiles: Group ${groupId} not found in groupMaps, attempting to resolve...`);
              const resolvedGroup = await resolveGroupInfo(groupId, deviceObjectId, userObjectId, token, groupMaps);
              
              if (resolvedGroup) {
                targetName = resolvedGroup.groupName;
                targetTypeInfo = state.targetMode === 'device' ? 'Device' : 'User';
                isMatch = true;
                matchCount++;
                logMessage(`pwshProfiles: MATCH - Resolved group ${targetName}`);
              } else {
                // Group exists in assignment but device/user is not a member - correctly filter out
                logMessage(`pwshProfiles: Group ${groupId} assignment filtered out (device/user not a member)`);
                continue;
              }
            }
          } else if (asg.target['@odata.type']) {
            const targetType = asg.target['@odata.type'].toLowerCase();
            if (targetType.includes('alldevicesassignmenttarget')) {
              targetName = 'All Devices';
              targetTypeInfo = 'Device';
              isMatch = true;
              matchCount++;
              logMessage(`pwshProfiles: MATCH - All Devices`);
            } else if (targetType.includes('allusersassignmenttarget') || targetType.includes('alllicensedusersassignmenttarget')) {
              targetName = 'All Users';
              targetTypeInfo = 'User';
              isMatch = true;
              matchCount++;
              logMessage(`pwshProfiles: MATCH - All Users`);
            } else {
              logMessage(`pwshProfiles: UNKNOWN target type ${targetType} - skipping`);
              continue;
            }
          } else {
            logMessage(`pwshProfiles: Assignment has no target info - skipping`);
            continue;
          }
          
          if (isMatch) {
            matchedScripts.push({
              scriptName: script.displayName,
              description: script.description || '',
              targetName,
              targetGroupId: asg.target.groupId || null, // Include group ID for dynamic group checking
              targetType: targetTypeInfo // Store the target type (Device/User)
            });
          }
        }
      }
      chrome.storage.local.set({ lastPwshAssignments: matchedScripts });
      updatePwshTable(matchedScripts);
      logMessage(`pwshProfiles: Found ${matchCount} matching assignments, saved ${matchedScripts.length} script entries`);
      if (matchCount === 0) {
        showResultNotification('No matching PowerShell profiles found for this device/user.', 'info');
      } else {
        showResultNotification(`PowerShell profiles loaded. Found ${matchCount} matches.`, 'success');
      }
    } catch (error) {
      logMessage(`pwshProfiles: Error - ${error.message}`);
      showResultNotification('Failed to load PowerShell profiles: ' + error.message, 'error');
    }
  };

  // Handle Group Creation
  const handleCreateGroup = async () => {
    logMessage("createGroup clicked");
    const groupName = document.getElementById("groupSearchInput").value.trim();
    if (!groupName) {
      showResultNotification('Please enter a group name.', 'error');
      return;
    }
    const mailNickname = groupName.substring(0, 10).replace(/[^a-zA-Z0-9]/g, '');
    chrome.storage.local.get("msGraphToken", async (data) => {
      if (!data.msGraphToken) {
        showResultNotification('No token captured. Please login first.', 'error');
        return;
      }
      try {
        const result = await fetchJSON('https://graph.microsoft.com/beta/groups', {
          method: 'POST',
          headers: {
            'Authorization': data.msGraphToken,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            displayName: groupName,
            mailEnabled: false,
            mailNickname: mailNickname,
            securityEnabled: true
          })
        });
        logMessage("Group created successfully");
        showResultNotification('Group created successfully!', 'success');
        // Refresh group list
        document.getElementById("searchGroup").click();
      } catch (error) {
        logMessage(`createGroup: Error - ${error.message}`);
        showResultNotification('Error creating group: ' + error.message, 'error');
      }
    });
  };

  // ── Collect Logs Modal Functions ───────────────────────────────────────
  
  // Supported environment variables for log paths
  const SUPPORTED_VARIABLES = [
    '%PROGRAMFILES%',
    '%PROGRAMDATA%',
    '%PUBLIC%',
    '%WINDIR%',
    '%TEMP%',
    '%TMP%'
  ];

  // Custom paths storage
  let customLogPaths = [];

  // Validate path starts with supported variable
  const validateLogPath = (path) => {
    const trimmedPath = path.trim();
    if (!trimmedPath) {
      return { valid: false, message: 'Path cannot be empty.' };
    }

    const upperPath = trimmedPath.toUpperCase();
    const hasValidVariable = SUPPORTED_VARIABLES.some(varName => 
      upperPath.startsWith(varName)
    );

    if (!hasValidVariable) {
      return { 
        valid: false, 
        message: `Path must start with one of: ${SUPPORTED_VARIABLES.join(', ')}` 
      };
    }

    return { valid: true };
  };

  // Show Collect Logs modal
  const showCollectLogsModal = () => {
    // Reset form
    document.getElementById('collectLogsAppId').value = '';
    document.getElementById('logPathPSADT').checked = false;
    document.getElementById('logPathIME').checked = true;
    customLogPaths = [];
    renderCustomPaths();
    clearInputError('customPathInput');
    clearInputError('collectLogsAppId');

    // Show modal
    document.getElementById('collectLogsModal').style.display = 'flex';
  };

  // Hide Collect Logs modal
  const hideCollectLogsModal = () => {
    document.getElementById('collectLogsModal').style.display = 'none';
  };

  // Add custom path to list
  const addCustomPath = () => {
    const input = document.getElementById('customPathInput');
    const path = input.value.trim();

    if (!path) return;

    const validation = validateLogPath(path);
    if (!validation.valid) {
      showInputError('customPathInput', validation.message);
      return;
    }

    // Check for duplicates
    if (customLogPaths.includes(path)) {
      showInputError('customPathInput', 'This path is already added.');
      return;
    }

    customLogPaths.push(path);
    input.value = '';
    clearInputError('customPathInput');
    renderCustomPaths();
  };

  // Remove custom path from list
  const removeCustomPath = (index) => {
    customLogPaths.splice(index, 1);
    renderCustomPaths();
  };

  // Render custom paths list
  const renderCustomPaths = () => {
    const container = document.getElementById('customPathsList');
    container.innerHTML = '';

    customLogPaths.forEach((path, index) => {
      const item = document.createElement('div');
      item.className = 'custom-path-item';
      
      const text = document.createElement('span');
      text.className = 'custom-path-text';
      text.textContent = path;
      
      const removeBtn = document.createElement('button');
      removeBtn.className = 'btn-remove-path';
      removeBtn.innerHTML = '<i class="material-icons">delete</i>';
      removeBtn.onclick = () => removeCustomPath(index);
      
      item.appendChild(text);
      item.appendChild(removeBtn);
      container.appendChild(item);
    });
  };

  // Show input error
  const showInputError = (inputId, message) => {
    const input = document.getElementById(inputId);
    input.classList.add('error');
    
    // Remove existing error message if any
    const existingError = input.parentElement.querySelector('.collect-logs-error-message');
    if (existingError) {
      existingError.remove();
    }
    
    // Add error message
    const errorMsg = document.createElement('p');
    errorMsg.className = 'collect-logs-error-message';
    errorMsg.innerHTML = `<i class="material-icons">error</i>${message}`;
    input.parentElement.appendChild(errorMsg);
  };

  // Clear input error
  const clearInputError = (inputId) => {
    const input = document.getElementById(inputId);
    input.classList.remove('error');
    
    const errorMsg = input.parentElement.querySelector('.collect-logs-error-message');
    if (errorMsg) {
      errorMsg.remove();
    }
  };

  // Gather all selected log paths and convert to semicolon-delimited format
  const gatherLogPaths = () => {
    const paths = [];

    // Get predefined paths
    const psadtCheckbox = document.getElementById('logPathPSADT');
    if (psadtCheckbox.checked) {
      paths.push(psadtCheckbox.getAttribute('data-path'));
    }

    const imeCheckbox = document.getElementById('logPathIME');
    if (imeCheckbox.checked) {
      paths.push(imeCheckbox.getAttribute('data-path'));
    }

    // Add custom paths
    paths.push(...customLogPaths);

    return paths;
  };

  // Handle Collecting Log Files
  const handleCollectLogs = () => {
    logMessage("collectLogs clicked");
    showCollectLogsModal();
  };

  // Confirm and execute log collection
  const confirmCollectLogs = async () => {
    logMessage("confirmCollectLogs clicked");

    // Validate app ID
    const appId = document.getElementById('collectLogsAppId').value.trim();
    if (!appId) {
      showInputError('collectLogsAppId', 'Application ID is required.');
      return;
    }
    clearInputError('collectLogsAppId');

    // Gather log paths
    const logPathsArray = gatherLogPaths();
    if (logPathsArray.length === 0) {
      showResultNotification('Please select or add at least one log path.', 'error');
      return;
    }

    // Convert to semicolon-delimited string
    const logPaths = logPathsArray.join(';');

    // Hide modal and proceed with log collection
    hideCollectLogsModal();

    try {
      const { mdmDeviceId } = await verifyMdmUrl();
      const token = await getToken();
      
      // Get device data to find the primary user
      const deviceData = await fetchJSON(`https://graph.microsoft.com/beta/deviceManagement/manageddevices('${mdmDeviceId}')`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });

      // Update device name display
      updateDeviceNameDisplay(deviceData);

      const userPrincipalName = deviceData.userPrincipalName;
      if (!userPrincipalName || userPrincipalName === 'Unknown user') {
        throw new Error("No primary user found for this device.");
      }

      // Get the user ID
      const userData = await fetchJSON(`https://graph.microsoft.com/beta/users?$filter=userPrincipalName eq '${encodeURIComponent(userPrincipalName)}'`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });

      if (!userData.value || userData.value.length === 0) {
        throw new Error("User not found in Azure AD.");
      }

      const userObjectId = userData.value[0].id;

      // Format the log paths for the API - with proper escaping
      const logPathsFormatted = [];
      logPaths.split(';').forEach(path => {
        const trimmedPath = path.trim();
        if (trimmedPath.length > 0) {
          // Replace single backslashes with double backslashes
          const escapedPath = trimmedPath.replace(/\\/g, '\\\\');
          logPathsFormatted.push(`"${escapedPath}"`);
        }
      });

      if (logPathsFormatted.length === 0) {
        throw new Error("No valid log paths provided.");
      }

      // Create raw JSON string with exactly the format needed
      const rawJsonBody = `{
        "customLogFolders": [${logPathsFormatted.join(', ')}],
        "id": "${userObjectId}_${mdmDeviceId}_${appId}"
      }`;

      logMessage(`collectLogs: Log paths: ${JSON.stringify(logPathsFormatted)}`);
      logMessage(`collectLogs: Requesting logs for user ${userObjectId}, device ${mdmDeviceId}, app ${appId}`);

      // Make the API call
      const requestUrl = `https://graph.microsoft.com/beta/users('${userObjectId}')/mobileAppTroubleshootingEvents('${mdmDeviceId}_${appId}')/appLogCollectionRequests`;

      const logResult = await fetchJSON(requestUrl, {
        method: "POST",
        headers: {
          "Authorization": token,
          "Content-Type": "application/json"
        },
        body: rawJsonBody
      });

      logMessage("collectLogs: Log collection request successful");
      showResultNotification("Log collection initiated successfully. The logs will be collected on the device and uploaded to Intune.", "success");
    } catch (error) {
      logMessage(`collectLogs: Error - ${error.message}`);
      showResultNotification('Failed to collect logs: ' + error.message, 'error');
    }
  };

  // ── Event Listener Registrations ───────────────────────────────────────
  document.getElementById('profileFilterInput').addEventListener('input', (e) => {
    const filterText = e.target.value.toLowerCase();
    filterTable(filterText);
    chrome.storage.local.set({
      profileFilterValue: filterText,
      currentDisplayType: state.currentDisplayType
    });
  });
  document.getElementById("searchGroup").addEventListener("click", handleSearchGroup);
  document.getElementById("addToGroups").addEventListener("click", handleAddToGroups);
  document.getElementById("removeFromGroups").addEventListener("click", handleRemoveFromGroups);
  document.getElementById("checkGroups").addEventListener("click", handleCheckGroups);
  document.getElementById("checkGroupMembers").addEventListener("click", handleCheckGroupMembers);
  document.getElementById("clearGroupMembers").addEventListener("click", handleClearGroupMembers);
  document.getElementById("checkGroupAssignments").addEventListener("click", handleCheckGroupAssignments);
  document.getElementById("checkCompliance").addEventListener("click", handleCheckCompliance);
  document.getElementById("downloadScript").addEventListener("click", handleDownloadScript);
  document.getElementById("appsAssignment").addEventListener("click", handleAppsAssignment);
  document.getElementById("pwshProfiles").addEventListener("click", handlePwshProfiles);
  document.getElementById("collectLogs").addEventListener("click", handleCollectLogs);
  document.getElementById("createGroup").addEventListener("click", handleCreateGroup);
  
  // Clear Members Modal Event Listeners
  document.getElementById("clearMembersModalClose").addEventListener("click", hideClearMembersModal);
  document.getElementById("clearSelectedMembersBtn").addEventListener("click", () => {
    showConfirmationSection('selected');
  });
  document.getElementById("clearAllMembersBtn").addEventListener("click", () => {
    showConfirmationSection('all');
  });
  document.getElementById("typedConfirmInput").addEventListener("input", (e) => {
    const isValid = e.target.value === 'REMOVE ALL';
    document.getElementById("confirmRemovalBtn").disabled = !isValid;
  });
  document.getElementById("confirmRemovalBtn").addEventListener("click", async () => {
    if (validateTypedConfirmation()) {
      await executeRemoval();
    }
  });
  document.getElementById("cancelRemovalBtn").addEventListener("click", () => {
    // Go back to action selection
    document.getElementById('confirmationSection').style.display = 'none';
    document.querySelector('.clear-action-buttons').style.display = 'flex';
    document.querySelector('.modal-description').style.display = 'block';
    clearMembersState.actionType = null;
  });
  document.getElementById("closeResultsBtn").addEventListener("click", hideClearMembersModal);
  
  // Collect Logs Modal Event Listeners
  document.getElementById("collectLogsModalClose").addEventListener("click", hideCollectLogsModal);
  document.getElementById("addCustomPathBtn").addEventListener("click", addCustomPath);
  document.getElementById("customPathInput").addEventListener("keypress", (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      addCustomPath();
    }
  });
  document.getElementById("confirmCollectLogsBtn").addEventListener("click", confirmCollectLogs);
  document.getElementById("cancelCollectLogsBtn").addEventListener("click", hideCollectLogsModal);
  
  // Close modal on Escape key
  document.addEventListener("keydown", (e) => {
    if (e.key === 'Escape') {
      const clearModal = document.getElementById('clearMembersModal');
      if (clearModal && clearModal.style.display === 'flex') {
        hideClearMembersModal();
      }
      
      const collectLogsModal = document.getElementById('collectLogsModal');
      if (collectLogsModal && collectLogsModal.style.display === 'flex') {
        hideCollectLogsModal();
      }
    }
  });
  
  // Close modal when clicking outside the dialog
  document.getElementById('clearMembersModal').addEventListener('click', (e) => {
    if (e.target.id === 'clearMembersModal') {
      hideClearMembersModal();
    }
  });
  
  document.getElementById('collectLogsModal').addEventListener('click', (e) => {
    if (e.target.id === 'collectLogsModal') {
      hideCollectLogsModal();
    }
  });
  
  document.getElementById("groupResults").addEventListener("change", (event) => {
    if (event.target.type === "checkbox") {
      // Clear table selections when selecting checkboxes
      clearTableSelection();
      updateActionButtonsState();

      chrome.storage.local.get(['lastSearchResults'], (data) => {
        if (data.lastSearchResults) {
          const updated = data.lastSearchResults.map(group =>
            group.id === event.target.value ? { ...group, checked: event.target.checked } : group
          );
          chrome.storage.local.set({ lastSearchResults: updated });
        }
      });
    }
  });
  // Use event delegation for sortable headers since they get replaced when table headers change
  document.addEventListener('click', function (e) {
    // Check if the clicked element is a sortable header
    if (e.target && e.target.classList.contains('sortable')) {
      state.sortDirection = state.sortDirection === 'asc' ? 'desc' : 'asc';
      e.target.classList.toggle('asc');
      e.target.classList.toggle('desc');
      // Re-render the current table based on display type
      if (state.currentDisplayType === 'config') {
        chrome.storage.local.get(['lastConfigAssignments'], (data) => {
          if (data.lastConfigAssignments) updateConfigTable(data.lastConfigAssignments, false);
        });
      } else if (state.currentDisplayType === 'apps') {
        chrome.storage.local.get(['lastAppAssignments'], (data) => {
          if (data.lastAppAssignments) updateAppTable(data.lastAppAssignments, false);
        });
      } else if (state.currentDisplayType === 'compliance') {
        chrome.storage.local.get(['lastComplianceAssignments'], (data) => {
          if (data.lastComplianceAssignments) updateComplianceTable(data.lastComplianceAssignments, false);
        });
      } else if (state.currentDisplayType === 'pwsh') {
        chrome.storage.local.get(['lastPwshAssignments'], (data) => {
          if (data.lastPwshAssignments) updatePwshTable(data.lastPwshAssignments, false);
        });
      } else if (state.currentDisplayType === 'groupMembers') {
        chrome.storage.local.get(['lastGroupMembers'], (data) => {
          if (data.lastGroupMembers) updateGroupMembersTable(data.lastGroupMembers, false);
        });
      } else if (state.currentDisplayType === 'groupAssignments') {
        chrome.storage.local.get(['lastGroupAssignments'], (data) => {
          if (data.lastGroupAssignments) updateGroupAssignmentsTable(data.lastGroupAssignments, false);
        });
      }
    }
  });
  // Theme toggle button
  document.getElementById("theme-toggle").addEventListener("click", toggleTheme);

  // Settings menu functionality
  const settingsButton = document.getElementById("settingsButton");
  const settingsDropdown = document.getElementById("settingsDropdown");
  
  settingsButton.addEventListener("click", (e) => {
    e.stopPropagation();
    settingsDropdown.classList.toggle("show");
  });

  // Close settings dropdown when clicking outside
  document.addEventListener("click", (e) => {
    if (!settingsButton.contains(e.target) && !settingsDropdown.contains(e.target)) {
      settingsDropdown.classList.remove("show");
    }
  });

  // Settings menu options
  document.getElementById("showWelcomeOption").addEventListener("click", (e) => {
    e.preventDefault();
    settingsDropdown.classList.remove("show");
    WelcomeNotification.showManual();
    logMessage('Welcome notification shown via settings menu');
  });

  document.getElementById("resetWelcomeOption").addEventListener("click", (e) => {
    e.preventDefault();
    settingsDropdown.classList.remove("show");
    WelcomeNotification.reset();
    showResultNotification('Welcome notification status reset - will show on next extension load', 'info');
    logMessage('Welcome notification status reset via settings menu');
  });

  // Clear Extension Storage option
  document.getElementById("clearStorageOption").addEventListener("click", (e) => {
    e.preventDefault();
    settingsDropdown.classList.remove("show");
    
    // Show confirmation dialog
    if (confirm('Are you sure you want to clear all extension storage? This will remove all cached data, settings, and search history. This action cannot be undone.')) {
      clearExtensionStorage();
    }
  });

  // Function to clear all extension storage
  const clearExtensionStorage = () => {
    // Preserve current theme before clearing storage
    const currentTheme = state.theme;
    
    chrome.storage.local.clear(() => {
      if (chrome.runtime.lastError) {
        showResultNotification('Error clearing extension storage: ' + chrome.runtime.lastError.message, 'error');
        logMessage('Error clearing extension storage: ' + chrome.runtime.lastError.message);
      } else {
        // Reset state variables to default values (but keep theme)
        state.currentDisplayType = 'config';
        state.sortDirection = 'asc';
        state.theme = currentTheme; // Keep current theme
        state.targetMode = 'device';
        state.selectedTableRows.clear();
        state.dynamicGroups.clear();
        state.pagination = {
          currentPage: 1,
          itemsPerPage: 10,
          totalItems: 0,
          totalPages: 0,
          filteredData: [],
          selectedRowIds: new Set()
        };

        // Restore theme setting to storage
        chrome.storage.local.set({ theme: currentTheme });

        // Reset UI to default state (but keep current theme)
        applyTheme(currentTheme);
        document.getElementById('deviceModeBtn').classList.add('active');
        document.getElementById('userModeBtn').classList.remove('active');
        document.getElementById('addBtnText').textContent = 'Add';
        document.getElementById('removeBtnText').textContent = 'Remove';
        
        // Clear all table content
        document.getElementById('configTableBody').innerHTML = '';
        document.getElementById('groupResults').innerHTML = '';
        document.getElementById('groupSearchInput').value = '';
        document.getElementById('profileFilterInput').value = '';
        
        // Reset pagination
        document.getElementById('paginationContainer').style.display = 'none';
        
        showResultNotification('Extension storage cleared successfully. All cached data has been reset (theme preference preserved).', 'success');
        logMessage('Extension storage cleared successfully via settings menu (theme preserved)');
      }
    });
  };

  // Pagination event listeners
  document.getElementById("prevPageBtn").addEventListener("click", () => {
    if (state.pagination.currentPage > 1) {
      goToPage(state.pagination.currentPage - 1);
    }
  });

  document.getElementById("nextPageBtn").addEventListener("click", () => {
    if (state.pagination.currentPage < state.pagination.totalPages) {
      goToPage(state.pagination.currentPage + 1);
    }
  });

  // Export CSV button event listener
  document.getElementById("exportCsvBtn").addEventListener("click", () => {
    exportTableToCsv();
  });

  // Keyboard navigation for pagination
  document.addEventListener('keydown', (e) => {
    // Only handle pagination keys when table is visible and has data
    if (state.pagination.totalItems > 0) {
      if (e.key === 'ArrowLeft' && e.ctrlKey && state.pagination.currentPage > 1) {
        e.preventDefault();
        goToPage(state.pagination.currentPage - 1);
      } else if (e.key === 'ArrowRight' && e.ctrlKey && state.pagination.currentPage < state.pagination.totalPages) {
        e.preventDefault();
        goToPage(state.pagination.currentPage + 1);
      } else if (e.key === 'Home' && e.ctrlKey && state.pagination.currentPage > 1) {
        e.preventDefault();
        goToPage(1);
      } else if (e.key === 'End' && e.ctrlKey && state.pagination.currentPage < state.pagination.totalPages) {
        e.preventDefault();
        goToPage(state.pagination.totalPages);
      }
    }
  });

  // Target mode toggle buttons
  document.getElementById("deviceModeBtn").addEventListener("click", () => handleTargetModeToggle('device'));
  document.getElementById("userModeBtn").addEventListener("click", () => handleTargetModeToggle('user'));

  // ── Welcome Notification Management ───────────────────────────────────
  // Add keyboard shortcut to show welcome notification (Ctrl+Shift+W)
  document.addEventListener('keydown', (e) => {
    if (e.ctrlKey && e.shiftKey && e.key === 'W') {
      e.preventDefault();
      WelcomeNotification.showManual();
      logMessage('Welcome notification shown manually via keyboard shortcut');
    }
    
    // Hidden admin shortcut to reset welcome status (Ctrl+Shift+Alt+R)
    if (e.ctrlKey && e.shiftKey && e.altKey && e.key === 'R') {
      e.preventDefault();
      WelcomeNotification.reset();
      showResultNotification('Welcome notification status reset - will show on next load', 'info');
      logMessage('Welcome notification status reset via keyboard shortcut');
    }
  });

  // Add help function for welcome notification
  window.showWelcome = () => {
    WelcomeNotification.showManual();
  };

  window.resetWelcome = () => {
    WelcomeNotification.reset();
  };

  // ── Initial Restoration Calls ─────────────────────────────────────────
  restoreState();
  restoreFilterValue();
  initializeTheme();
  
  // Set version number in settings dropdown
  const versionElement = document.getElementById('extensionVersion');
  if (versionElement && chrome.runtime && chrome.runtime.getManifest) {
    versionElement.textContent = chrome.runtime.getManifest().version;
  }
});
