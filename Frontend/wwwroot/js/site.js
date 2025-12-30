$(document).ready(function () {
    const ApiUrl = 'https://localhost:7086';

    console.log("API URL : ", ApiUrl)

    // Download Excel file
    $('#BtnView').click(function () {
        downloadExcelFile();
    });

    // Preview Excel file
    $('#BtnPreview').click(function () {
        previewExcelFile();
    });

    function downloadExcelFile() {
        // Get temp file path first
        $.ajax({
            url: ApiUrl + "/api/excel/temp-path",
            type: 'GET',
            success: function (response) {
                console.log("Response : ", response);
                // Download using the temp file name
                const downloadUrl = ApiUrl +`/api/excel/download-temp/${response.fileName}`;

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

    function previewExcelFile() {
        // Show modal
        $('#excelPreviewModal').modal('show');

        // Show loading
        $('#loadingSpinner').show();
        $('#excelPreviewContainer').hide();
        $('#errorMessage').hide();

        // Get temp file path first
        $.ajax({
            url: ApiUrl + `/api/excel/temp-path`,
            type: 'GET',
            success: function (response) {
                console.log("Response : ", response);
                console.log('Temp file created for preview:', response.fileName);

                // Now get the file content for preview using the temp endpoint
                $.ajax({
                    url: ApiUrl +`/api/excel/download-temp/${response.fileName}`,
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
                                $('#excelTableContainer').html(htmlTable);
                                $('#excelTableContainer table').addClass('table table-striped table-bordered table-sm');

                                $('#loadingSpinner').hide();
                                $('#excelPreviewContainer').show();

                            } catch (error) {
                                showError('Failed to parse Excel file: ' + error.message);
                            }
                        };

                        fileReader.readAsArrayBuffer(data);
                    },
                    error: function (xhr, status, error) {
                        console.error('Error downloading temp file:', error);
                        showError('Failed to load Excel file for preview.');
                    }
                });
            },
            error: function (xhr, status, error) {
                console.error('Error generating temp file:', error);
                showError('Failed to generate Excel file for preview.');
            }
        });
    }

    function showError(message) {
        $('#loadingSpinner').hide();
        $('#excelPreviewContainer').hide();
        $('#errorMessage').text(message).show();
    }
});