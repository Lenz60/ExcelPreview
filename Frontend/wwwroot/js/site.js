// Please see documentation at https://docs.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.
$(document).ready(function () {

    $('#BtnView').click(function () {
        // Get the API URL from ViewBag or use a fallback
        const apiUrl = window.apiUrl || 'https://localhost:7086';

        //console.log('Button clicked');

        $.ajax({
            url: `${apiUrl}/api/excel/download`,
            type: 'GET',
            xhrFields: {
                responseType: 'blob' // Important for binary data
            },
            success: function (data, status, xhr) {
                // Get filename from Content-Disposition header or use default
                const contentDisposition = xhr.getResponseHeader('Content-Disposition');
                let filename = 'GeneratedExcelFile.xlsx';

                if (contentDisposition) {
                    const filenameMatch = contentDisposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
                    if (filenameMatch && filenameMatch[1]) {
                        filename = filenameMatch[1].replace(/['"]/g, '');
                    }
                }

                // Create blob and download
                const blob = new Blob([data], {
                    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                });

                const link = document.createElement('a');
                link.href = window.URL.createObjectURL(blob);
                link.download = filename;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);

                // Clean up the blob URL
                window.URL.revokeObjectURL(link.href);

                console.log('Excel file downloaded successfully');
            },
            error: function (xhr, status, error) {
                console.error('Error generating Excel file:', error);
                console.error('Status:', status);
                console.error('Response:', xhr.responseText);
                alert('Failed to download Excel file. Please try again.');
            }
        });
    });
});