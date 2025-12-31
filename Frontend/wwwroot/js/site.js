$(document).ready(function () {
    const ApiUrl = 'https://localhost:7086';
    let isExcel = true; // Fix: use let instead of const for reassignment

    console.log("API URL : ", ApiUrl)

    // Download Excel file
    $('#BtnView').click(function () {
        isExcel = true;
        downloadExcelFile();
    });

    // Preview Excel file
    $('#BtnPreview').click(function () {
        isExcel = true;
        previewFile();
    });

    // Download Pdf file
    $('#BtnViewPdf').click(function () {
        isExcel = false;
        downloadPdfFile();
    });

    // Preview Pdf file
    $('#BtnPreviewPdf').click(function () {
        isExcel = false;
        previewFile();
    });

    function downloadExcelFile() {
        // Get temp file path first
        $.ajax({
            url: ApiUrl + "/api/excel/temp-path",
            type: 'GET',
            success: function (response) {
                console.log("Response : ", response);
                // Download using the temp file name
                const downloadUrl = ApiUrl + `/api/excel/download-temp/${response.fileName}`;

                // Create invisible download link
                const link = document.createElement('a');
                link.href = downloadUrl;
                link.download = response.fileName;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);

                console.log('Excel file download initiated:', response.fileName);
            },
            error: function (xhr, status, error) {
                console.error('Error generating Excel file:', error);
                alert('Failed to generate Excel file. Please try again.');
            }
        });
    }

    function downloadPdfFile() {
        // Get temp file path first
        $.ajax({
            url: ApiUrl + "/api/excel/pdf-temp-path",
            type: 'GET',
            success: function (response) {
                console.log("Response : ", response);
                // Download using the temp file name
                const downloadUrl = ApiUrl + `/api/excel/download-pdf-temp/${response.fileName}`;

                // Create invisible download link
                const link = document.createElement('a');
                link.href = downloadUrl;
                link.download = response.fileName;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);

                console.log('PDF file download initiated:', response.fileName);
            },
            error: function (xhr, status, error) {
                console.error('Error generating PDF file:', error);
                alert('Failed to generate PDF file. Please try again.');
            }
        });
    }

    function previewFile() {
        // Update modal title and loading message based on context
        if (isExcel) {
            $('#previewModal .modal-title').text('Excel Preview');
            $('#loadingMessage').text('Loading Excel preview...');
            previewExcelFile();
        } else {
            $('#previewModal .modal-title').text('PDF Preview');
            $('#loadingMessage').text('Loading PDF preview...');
            previewPdfFile();
        }

        // Show modal
        $('#previewModal').modal('show');

        // Show loading
        $('#loadingSpinner').show();
        $('#previewContainer').hide();
        $('#errorMessage').hide();
    }

    function previewPdfFile() {
        // Get temp file path first
        $.ajax({
            url: ApiUrl + `/api/excel/pdf-temp-path`,
            type: 'GET',
            success: function (response) {
                console.log("PDF Response: ", response);
                console.log('Temp PDF file created for preview:', response.fileName);

                // Use the PREVIEW endpoint instead of download endpoint
                const pdfPreviewUrl = ApiUrl + `/api/excel/preview-pdf-temp/${response.fileName}`;

                // Method 1: Using iframe (most compatible)
                const pdfEmbed = `
                    <div class="pdf-container" style="width: 100%; height: 700px; border: 1px solid #ddd;">
                        <iframe src="${pdfPreviewUrl}" 
                                width="100%" 
                                height="100%" 
                                frameborder="0" 
                                style="border: none;">
                            <div class="alert alert-info text-center">
                                <p><i class="fas fa-exclamation-triangle"></i> PDF preview is not supported in your browser.</p>
                                <p><a href="${pdfPreviewUrl}" target="_blank" class="btn btn-primary">Open PDF in new tab</a></p>
                            </div>
                        </iframe>
                    </div>
                `;

                // Method 2: Alternative using object tag (uncomment if needed)
                /*
                const pdfEmbed = `
                    <div class="pdf-container" style="width: 100%; height: 700px;">
                        <object data="${pdfPreviewUrl}" type="application/pdf" width="100%" height="100%">
                            <div class="alert alert-info text-center">
                                <p><i class="fas fa-exclamation-triangle"></i> PDF preview is not supported in your browser.</p>
                                <p><a href="${pdfPreviewUrl}" target="_blank" class="btn btn-primary">Open PDF in new tab</a></p>
                            </div>
                        </object>
                    </div>
                `;
                */

                // Display PDF in modal
                $('#fileContentContainer').html(pdfEmbed);

                $('#loadingSpinner').hide();
                $('#previewContainer').show();

            },
            error: function (xhr, status, error) {
                console.error('Error generating PDF temp file:', error);
                showError('Failed to generate PDF file for preview.');
            }
        });
    }

    function previewExcelFile() {
        // Get temp file path first
        $.ajax({
            url: ApiUrl + `/api/excel/temp-path`,
            type: 'GET',
            success: function (response) {
                console.log("Excel Response: ", response);
                console.log('Temp Excel file created for preview:', response.fileName);

                // Now get the file content for preview using the temp endpoint
                $.ajax({
                    url: ApiUrl + `/api/excel/download-temp/${response.fileName}`,
                    type: 'GET',
                    xhrFields: {
                        responseType: 'blob'
                    },
                    success: function (data) {
                        const fileReader = new FileReader();
                        fileReader.onload = function (e) {
                            try {
                                const arrayBuffer = e.target.result;
                                const workbook = XLSX.read(arrayBuffer, { type: 'array' });

                                // Get first sheet
                                const firstSheetName = workbook.SheetNames[0];
                                const worksheet = workbook.Sheets[firstSheetName];

                                // Convert to HTML
                                const htmlTable = XLSX.utils.sheet_to_html(worksheet);

                                // Display in modal
                                $('#fileContentContainer').html(`
                                    <div class="table-responsive">
                                        ${htmlTable}
                                    </div>
                                `);
                                $('#fileContentContainer table').addClass('table table-striped table-bordered table-sm');

                                $('#loadingSpinner').hide();
                                $('#previewContainer').show();

                            } catch (error) {
                                showError('Failed to parse Excel file: ' + error.message);
                            }
                        };

                        fileReader.readAsArrayBuffer(data);
                    },
                    error: function (xhr, status, error) {
                        console.error('Error downloading temp Excel file:', error);
                        showError('Failed to load Excel file for preview.');
                    }
                });
            },
            error: function (xhr, status, error) {
                console.error('Error generating Excel temp file:', error);
                showError('Failed to generate Excel file for preview.');
            }
        });
    }

    function showError(message) {
        $('#loadingSpinner').hide();
        $('#previewContainer').hide();
        $('#errorMessage').text(message).show();
    }
});