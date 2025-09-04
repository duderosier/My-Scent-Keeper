/**
 * Advanced Backup Manager for My Scent Keeper
 * Handles Excel export/import with optimized image processing
 */

class BackupManager {
    constructor() {
        this.maxImageSize = 1024 * 1024; // 1MB limit for Excel images
        this.compressionQuality = 0.7;
        this.batchSize = 10; // Process items in batches to avoid memory issues
    }

    /**
     * Export inventory to Excel with optimized image handling
     */
    async exportToExcel(inventory, options = {}) {
        const {
            includeImages = true,
            filename = null,
            onProgress = null
        } = options;

        try {
            if (!inventory || inventory.length === 0) {
                throw new Error('No data to export');
            }

            const totalSteps = inventory.length + 2; // +2 for setup and finalization
            let currentStep = 0;

            const updateProgress = (message) => {
                currentStep++;
                if (onProgress) {
                    onProgress({
                        step: currentStep,
                        total: totalSteps,
                        message,
                        percentage: Math.round((currentStep / totalSteps) * 100)
                    });
                }
            };

            updateProgress('Initializing export...');

            // Prepare data structure
            const data = [];
            const headers = [
                'Product Name', 'Barcode', 'Quantity', 'Location', 
                'Rating', 'Notes', 'Date Added'
            ];

            if (includeImages) {
                headers.push('Image Data', 'Image Type');
            }

            data.push(headers);

            // Process inventory in batches
            for (let i = 0; i < inventory.length; i += this.batchSize) {
                const batch = inventory.slice(i, i + this.batchSize);
                
                for (const item of batch) {
                    updateProgress(`Processing ${item.name}...`);
                    
                    const row = [
                        item.name || '',
                        item.barcode || '',
                        item.quantity || 0,
                        item.location || '',
                        item.rating || 0,
                        item.notes || '',
                        item.dateAdded ? new Date(item.dateAdded).toLocaleDateString() : ''
                    ];

                    if (includeImages) {
                        try {
                            const imageData = await this.getOptimizedImageData(item);
                            row.push(imageData.data || '');
                            row.push(imageData.type || '');
                        } catch (imageError) {
                            console.warn('Image processing failed for:', item.name, imageError);
                            row.push('', ''); // Empty image data
                        }
                    }

                    data.push(row);
                }

                // Allow UI to update between batches
                await this.delay(10);
            }

            updateProgress('Creating Excel file...');

            // Create workbook
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(data);

            // Set column widths
            const colWidths = [
                { width: 25 }, // Product Name
                { width: 15 }, // Barcode
                { width: 10 }, // Quantity
                { width: 20 }, // Location
                { width: 10 }, // Rating
                { width: 30 }, // Notes
                { width: 15 }  // Date Added
            ];

            if (includeImages) {
                colWidths.push({ width: 20 }); // Image Data
                colWidths.push({ width: 15 }); // Image Type
            }

            ws['!cols'] = colWidths;

            XLSX.utils.book_append_sheet(wb, ws, 'Inventory');

            // Generate filename
            const timestamp = new Date().toISOString().slice(0, 19).replace(/[:.]/g, '-');
            const exportFilename = filename || `my-scent-keeper-backup-${timestamp}.xlsx`;

            updateProgress('Downloading file...');

            // Download file
            XLSX.writeFile(wb, exportFilename);

            return {
                success: true,
                filename: exportFilename,
                itemCount: inventory.length,
                includeImages
            };

        } catch (error) {
            console.error('Export failed:', error);
            throw new Error(`Export failed: ${error.message}`);
        }
    }

    /**
     * Import inventory from Excel file
     */
    async importFromExcel(file, options = {}) {
        const {
            replaceExisting = true,
            onProgress = null,
            loadImages = true
        } = options;

        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = async (e) => {
                try {
                    if (onProgress) onProgress({ message: 'Reading file...', percentage: 10 });

                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                    if (jsonData.length < 2) {
                        throw new Error('Invalid file format. No data found.');
                    }

                    if (onProgress) onProgress({ message: 'Processing data...', percentage: 30 });

                    const headers = jsonData[0];
                    const hasImages = headers.includes('Image Data');
                    const importedInventory = [];

                    // Process data rows
                    for (let i = 1; i < jsonData.length; i++) {
                        const row = jsonData[i];
                        
                        if (row[0]) { // Only process rows with product names
                            const product = {
                                id: Date.now() + i + Math.random() * 1000,
                                name: row[0] || '',
                                barcode: row[1] || 'N/A',
                                quantity: parseInt(row[2]) || 1,
                                location: row[3] || 'Unspecified',
                                rating: parseInt(row[4]) || 0,
                                notes: row[5] || '',
                                dateAdded: row[6] ? new Date(row[6]).toISOString() : new Date().toISOString()
                            };

                            // Handle image data if present
                            if (hasImages && loadImages && row[7]) {
                                try {
                                    product.image = row[7];
                                    product.imageType = row[8] || 'image/jpeg';
                                } catch (imageError) {
                                    console.warn('Failed to process image for:', product.name, imageError);
                                }
                            }

                            importedInventory.push(product);
                        }

                        // Update progress
                        if (onProgress) {
                            const percentage = 30 + ((i / jsonData.length) * 60);
                            onProgress({ 
                                message: `Processing item ${i} of ${jsonData.length - 1}...`, 
                                percentage: Math.round(percentage) 
                            });
                        }
                    }

                    if (onProgress) onProgress({ message: 'Import completed!', percentage: 100 });

                    resolve({
                        success: true,
                        inventory: importedInventory,
                        itemCount: importedInventory.length,
                        hasImages: hasImages && loadImages
                    });

                } catch (error) {
                    reject(new Error(`Import failed: ${error.message}`));
                }
            };

            reader.onerror = () => {
                reject(new Error('Failed to read file'));
            };

            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * Get optimized image data for Excel export
     */
    async getOptimizedImageData(item) {
        try {
            // Try to get image from IndexedDB first, then fallback to item.image
            let imageData = null;
            
            if (typeof loadProductImage === 'function') {
                imageData = await loadProductImage(item.id);
            }
            
            if (!imageData && item.image) {
                imageData = item.image;
            }

            if (!imageData) {
                return { data: '', type: '' };
            }

            // Check if image needs compression
            if (this.getDataUrlSize(imageData) > this.maxImageSize) {
                return await this.compressImageForExcel(imageData);
            }

            return {
                data: imageData,
                type: this.getImageTypeFromDataUrl(imageData)
            };

        } catch (error) {
            console.warn('Image optimization failed:', error);
            return { data: '', type: '' };
        }
    }

    /**
     * Compress image specifically for Excel export
     */
    async compressImageForExcel(dataUrl) {
        return new Promise((resolve) => {
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            const img = new Image();

            img.onload = () => {
                // Calculate dimensions (max 800px for Excel)
                const maxSize = 800;
                let { width, height } = img;

                if (width > height) {
                    if (width > maxSize) {
                        height = height * (maxSize / width);
                        width = maxSize;
                    }
                } else {
                    if (height > maxSize) {
                        width = width * (maxSize / height);
                        height = maxSize;
                    }
                }

                canvas.width = width;
                canvas.height = height;

                // Draw and compress
                ctx.drawImage(img, 0, 0, width, height);
                
                const compressedDataUrl = canvas.toDataURL('image/jpeg', this.compressionQuality);
                
                resolve({
                    data: compressedDataUrl,
                    type: 'image/jpeg'
                });
            };

            img.onerror = () => {
                resolve({ data: '', type: '' });
            };

            img.src = dataUrl;
        });
    }

    /**
     * Utility functions
     */
    getDataUrlSize(dataUrl) {
        if (!dataUrl) return 0;
        // Rough estimation: data URL size in bytes
        return dataUrl.length * 0.75;
    }

    getImageTypeFromDataUrl(dataUrl) {
        if (!dataUrl) return '';
        const match = dataUrl.match(/^data:([^;]+);/);
        return match ? match[1] : 'image/jpeg';
    }

    delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    /**
     * Create progress modal for better user experience
     */
    createProgressModal() {
        const modal = document.createElement('div');
        modal.className = 'backup-progress-modal';
        modal.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.7);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 9999;
        `;

        const content = document.createElement('div');
        content.style.cssText = `
            background: white;
            padding: 30px;
            border-radius: 15px;
            text-align: center;
            max-width: 400px;
            width: 90%;
        `;

        const title = document.createElement('h3');
        title.textContent = 'Processing Backup...';
        title.style.marginBottom = '20px';

        const progressBar = document.createElement('div');
        progressBar.style.cssText = `
            width: 100%;
            height: 20px;
            background: #f0f0f0;
            border-radius: 10px;
            overflow: hidden;
            margin-bottom: 15px;
        `;

        const progressFill = document.createElement('div');
        progressFill.style.cssText = `
            height: 100%;
            background: linear-gradient(45deg, #8b5cf6, #7c3aed);
            width: 0%;
            transition: width 0.3s ease;
        `;

        const statusText = document.createElement('div');
        statusText.style.cssText = `
            color: #666;
            font-size: 14px;
        `;

        progressBar.appendChild(progressFill);
        content.appendChild(title);
        content.appendChild(progressBar);
        content.appendChild(statusText);
        modal.appendChild(content);

        return {
            modal,
            update: (progress) => {
                progressFill.style.width = `${progress.percentage || 0}%`;
                statusText.textContent = progress.message || '';
            },
            remove: () => {
                if (modal.parentNode) {
                    modal.parentNode.removeChild(modal);
                }
            }
        };
    }
}

// Initialize global backup manager
window.backupManager = new BackupManager();

// Enhanced export function
async function exportToExcelAdvanced(includeImages = true) {
    const progressModal = backupManager.createProgressModal();
    document.body.appendChild(progressModal.modal);

    try {
        const result = await backupManager.exportToExcel(inventory, {
            includeImages,
            onProgress: progressModal.update
        });

        progressModal.remove();
        
        const successDiv = document.createElement('div');
        successDiv.className = 'loading-indicator';
        successDiv.style.background = '#d4edda';
        successDiv.style.color = '#155724';
        successDiv.textContent = `Export successful! Downloaded ${result.itemCount} items${result.includeImages ? ' with images' : ''}.`;
        document.getElementById('backup').appendChild(successDiv);

        setTimeout(() => successDiv.remove(), 3000);

    } catch (error) {
        progressModal.remove();
        console.error('Export failed:', error);
        alert(`Export failed: ${error.message}`);
    }
}

// Enhanced import function
async function importFromExcelAdvanced(event) {
    const file = event.target.files[0];
    if (!file) return;

    const progressModal = backupManager.createProgressModal();
    document.body.appendChild(progressModal.modal);

    try {
        const result = await backupManager.importFromExcel(file, {
            onProgress: progressModal.update,
            loadImages: true
        });

        progressModal.remove();

        if (confirm(`Import ${result.itemCount} items${result.hasImages ? ' with images' : ''}? This will replace your current inventory.`)) {
            inventory = result.inventory;
            await saveInventory();
            await updateInventoryDisplay();
            updateLocationFilter();
            
            alert(`Successfully imported ${result.itemCount} products!`);
        }

    } catch (error) {
        progressModal.remove();
        console.error('Import failed:', error);
        alert(`Import failed: ${error.message}`);
    }

    // Reset file input
    event.target.value = '';