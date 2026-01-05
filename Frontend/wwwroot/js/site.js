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

    $.ajaxSetup({
        xhrFields: {
            withCredentials: true  // This is CRITICAL for cross-origin cookies
        },
        crossDomain: true
    });


    function getCookie(name) {
        return document.cookie.split('; ').reduce((r, v) => {
            const parts = v.split('=');
            return parts[0] === name ? decodeURIComponent(parts[1]) : r;
        }, '');
    }

    function downloadExcelFile() {
        // Get temp file path first
        $.ajax({
            url: ApiUrl + "/api/excel/temp-path",
            type: 'GET',
            success: function (response) {
                console.log("Response : ", response);
                // Download using the temp file name
                const downloadUrl = ApiUrl + `/api/excel/download-temp/`;

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
                const downloadUrl = ApiUrl + `/api/excel/download-pdf-temp`;

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
        const tempFileName = getCookie('tempPdfFileName');
        if (!tempFileName) {
            $.ajax({
                url: ApiUrl + `/api/excel/pdf-temp-path`,
                type: 'GET',
                success: function (response) {
                    console.log("PDF Response: ", response);
                    console.log('Temp PDF file created for preview:', response.fileName);

                    loadPdfPreview(response.fileName);

                },
                error: function (xhr, status, error) {
                    console.error('Error generating PDF temp file:', error);
                    showError('Failed to generate PDF file for preview.');
                }
            });

        } else {
            loadPdfPreview(tempFileName);
            console.log('Temp PDF file get from cookie:', tempFileName);
        }
    }

    function loadExcelPreview(fileName) {
        $.ajax({
            url: ApiUrl + `/api/excel/preview-temp/${fileName}`, // 🔑 Use preview endpoint
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
                console.error('Error previewing temp Excel file:', error);
                showError('Failed to load Excel file for preview.');
            }
        });
    }

    function loadPdfPreview(fileName) {
        $.ajax({
            url: ApiUrl + `/api/excel/preview-pdf-temp/${fileName}`,
            type: 'GET',
            xhrFields: {
                responseType: 'blob',
                withCredentials: true
            },
            success: function (blob) {
                console.log('✅ PDF blob received, size:', blob.size);

                // Create blob URL for the PDF
                const blobUrl = URL.createObjectURL(blob);
                console.log('🔗 Blob URL created:', blobUrl);

                const pdfEmbed = `
                    <div class="pdf-container" style="width: 100%; height: 700px; border: 1px solid #ddd;">
                        <iframe src="${blobUrl}" 
                                width="100%" 
                                height="100%" 
                                frameborder="0" 
                                style="border: none;">
                            <div class="alert alert-info text-center">
                                <p>PDF preview not supported. <a href="${blobUrl}" target="_blank" class="btn btn-primary">Open PDF</a></p>
                            </div>
                        </iframe>
                    </div>
                `;

                $('#fileContentContainer').html(pdfEmbed);
                $('#loadingSpinner').hide();
                $('#previewContainer').show();

                // Clean up blob URL after 5 minutes
                setTimeout(() => {
                    URL.revokeObjectURL(blobUrl);
                    console.log('🧹 Blob URL cleaned up');
                }, 300000);
            },
            error: function (xhr, status, error) {
                console.error('❌ Error loading PDF blob:', error);
                console.error('Status:', xhr.status, 'Response:', xhr.responseText);
                showError(`Failed to load PDF file: ${error}`);
            }
        });
    }

    function previewExcelFile() {
        // Get temp file path first
        const tempFileName = getCookie('tempExcelFileName');
        console.log("Temp File Name from cookie: ", tempFileName);
        if (!tempFileName) {
            $.ajax({
                url: ApiUrl + `/api/excel/temp-path`,
                type: 'GET',
                success: function (response) {
                    console.log("Excel Response: ", response);
                    console.log('Temp Excel file created for preview:', response.fileName);
                    //localStorage.setItem('tempExcelFileName', response.fileName);

                    // Now get the file content for preview using the temp endpoint
                    loadExcelPreview(response.fileName);
                },
                error: function (xhr, status, error) {
                    console.error('Error generating Excel temp file:', error);
                    showError('Failed to generate Excel file for preview.');
                }
            });
        } else {
            console.log('Temp Excel file get from localstorage:', tempFileName);
            loadExcelPreview(tempFileName);

        }
    }

    function showError(message) {
        $('#loadingSpinner').hide();
        $('#previewContainer').hide();
        $('#errorMessage').text(message).show();
    }
});