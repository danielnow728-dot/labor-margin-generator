document.addEventListener('DOMContentLoaded', () => {
    // State management
    const files = {
        gl: null,
        inventory: null,
        cost: null,
        master: null
    };

    const dropZones = document.querySelectorAll('.drop-zone');
    const generateBtn = document.getElementById('generate-btn');
    const statusIndicator = document.getElementById('status-indicator');
    const statusText = document.getElementById('status-text');
    const spinner = document.getElementById('spinner');

    // UI Toast System
    function showToast(message, type = 'success') {
        const toast = document.getElementById('toast');
        const icon = document.getElementById('toast-icon');
        const text = document.getElementById('toast-message');

        text.textContent = message;
        toast.className = `toast show ${type}`;

        icon.className = type === 'success'
            ? 'ri-checkbox-circle-fill'
            : 'ri-error-warning-fill';

        setTimeout(() => {
            toast.className = 'toast hidden';
        }, 4000);
    }

    // Check if the user has uploaded all required files
    function checkReadyState() {
        const missingCount = Object.values(files).filter(f => !f).length;

        if (missingCount === 0) {
            generateBtn.disabled = false;
            statusText.textContent = "All files loaded. Ready to generate.";
            statusIndicator.className = "status-indicator ready";
        } else {
            generateBtn.disabled = true;
            statusText.textContent = `Waiting for ${missingCount} file${missingCount > 1 ? 's' : ''}...`;
            statusIndicator.className = "status-indicator";
        }
    }

    // Handle a file being attached to a drop zone
    function handleFileSelection(file, type, zone) {
        if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
            showToast('Please upload an Excel file (.xlsx or .xls)', 'error');
            return;
        }

        files[type] = file;

        // Update UI
        zone.classList.add('has-file');
        const nameEl = zone.querySelector('.file-name');
        nameEl.textContent = file.name;

        checkReadyState();
    }

    // Setup Drag, Drop, and Click on all zones
    dropZones.forEach(zone => {
        const type = zone.closest('.upload-card').dataset.fileType;
        const input = zone.querySelector('input[type="file"]');

        // Click to browse
        zone.addEventListener('click', () => {
            input.click();
        });

        // Handle file browse selection
        input.addEventListener('change', (e) => {
            if (e.target.files.length) {
                handleFileSelection(e.target.files[0], type, zone);
            }
        });

        // Drag events
        zone.addEventListener('dragover', (e) => {
            e.preventDefault();
            zone.classList.add('dragover');
        });

        zone.addEventListener('dragleave', () => {
            zone.classList.remove('dragover');
        });

        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            zone.classList.remove('dragover');

            if (e.dataTransfer.files.length) {
                handleFileSelection(e.dataTransfer.files[0], type, zone);
            }
        });
    });

    // Handle Generation Submission
    generateBtn.addEventListener('click', async () => {
        // Validation check (though button is disabled, just in case)
        if (Object.values(files).some(f => !f)) return;

        // UI Loading State
        generateBtn.disabled = true;
        generateBtn.querySelector('.btn-text').textContent = "Processing...";
        spinner.classList.remove('hidden');
        statusIndicator.className = "status-indicator processing";
        statusText.textContent = "Crunching thousands of rows, please wait...";

        // Prepare forms data
        const formData = new FormData();
        formData.append('gl', files.gl);
        formData.append('inventory', files.inventory);
        formData.append('cost', files.cost);
        formData.append('master', files.master);

        try {
            // Note: Update URL if running backend on a different port/host
            const response = await fetch('http://localhost:8000/generate-report', {
                method: 'POST',
                body: formData,
            });

            if (!response.ok) {
                const errData = await response.json().catch(() => ({}));
                throw new Error(errData.detail || `Server error: ${response.status}`);
            }

            // Successfully received the file back
            const blob = await response.blob();
            const downloadUrl = window.URL.createObjectURL(blob);

            // Create a temporary link to trigger the download
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = downloadUrl;

            // Look for disposition header or default name
            const disposition = response.headers.get('content-disposition');
            let filename = "Labor_Margin_Report.xlsx";
            if (disposition && disposition.indexOf('attachment') !== -1) {
                const matches = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/.exec(disposition);
                if (matches != null && matches[1]) {
                    filename = matches[1].replace(/['"]/g, '');
                }
            }

            a.download = filename;
            document.body.appendChild(a);
            a.click();

            // Cleanup UI
            window.URL.revokeObjectURL(downloadUrl);
            a.remove();
            showToast('Success! File downloaded to your computer.', 'success');

        } catch (error) {
            console.error('Generation Error:', error);
            showToast(`Error: ${error.message}`, 'error');
        } finally {
            // Reset UI Loading State
            generateBtn.disabled = false;
            generateBtn.querySelector('.btn-text').textContent = "Generate Report";
            spinner.classList.add('hidden');
            checkReadyState(); // Revert status correctly
        }
    });

    // Global drag-drop overrides to prevent browser from taking over screen
    window.addEventListener("dragover", function (e) {
        e.preventDefault();
    });
    window.addEventListener("drop", function (e) {
        e.preventDefault();
    });
});
