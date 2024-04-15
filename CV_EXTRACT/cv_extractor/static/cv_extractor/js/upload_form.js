function updateProgress(percentage, text) {
    document.getElementById('progress-bar').value = percentage;
    document.getElementById('progress-percentage').innerText = text; // Update percentage text
    document.getElementById('progress-bar-container').style.display = 'block';
}

function showDownloadButton() {
    document.getElementById('download-button').style.display = 'block';
}

function hideDownloadButton() {
    document.getElementById('download-button').style.display = 'none';
}

document.getElementById('document').onchange = function() {
    // Hide the download button when a new file is selected
    hideDownloadButton();
    // Reset progress bar and text
    updateProgress(0, '0% - Ready to upload');
};

document.getElementById('upload-form').onsubmit = function(event) {
    event.preventDefault();
    var formData = new FormData(this);
    var xhr = new XMLHttpRequest();
    xhr.open('POST', this.action, true);

    xhr.upload.onprogress = function(e) {
        if (e.lengthComputable) {
            var uploadPercentage = (e.loaded / e.total) * 100;
            updateProgress(uploadPercentage, uploadPercentage.toFixed(0) + '% - Uploading');
        }
    };

    xhr.onloadstart = function() {
        // Indicate that processing has started when the load starts
        updateProgress(0, '0% - Starting processing');
    };

    xhr.onload = function() {
        if (this.status == 200) {
            var response = JSON.parse(this.response);
            window.file_path = response.file_path; // Save the file path
            updateProgress(100, '100% - Ready to download');
            showDownloadButton(); // Show download button
        } else {
            // Handle error
            document.getElementById('progress-bar-container').style.display = 'none';
            alert('An error occurred while processing the file. Please try again.');
        }
    };

    xhr.send(formData);
};

function downloadFile() {
var downloadUrl = '/cv/download/' + encodeURIComponent(window.file_path) + '/';
window.location.href = downloadUrl; // Use the saved file path to download the file
hideDownloadButton(); // Hide download button after initiating download
}