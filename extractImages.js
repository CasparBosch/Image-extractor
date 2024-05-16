document.getElementById('docxFile').addEventListener('change', handleFileSelect, false);

let selectedFile;

function handleFileSelect(event) {
    const file = event.target.files[0];
    if (file && file.name.endsWith('.docx')) {
        selectedFile = file;
    } else {
        alert('Please select a valid .docx file.');
    }
}

async function extractImages() {
    if (!selectedFile) {
        alert('Please select a .docx file first.');
        return;
    }

    try {
        const arrayBuffer = await selectedFile.arrayBuffer();
        const zip = await JSZip.loadAsync(arrayBuffer);
        const imgFolder = zip.folder('word/media');

        if (!imgFolder) {
            alert('No images found in the document.');
            return;
        }

        const imagesContainer = document.getElementById('imagesContainer');
        imagesContainer.innerHTML = '';

        imgFolder.forEach(async (relativePath, file) => {
            const blob = await file.async('blob');
            const imgElement = document.createElement('img');
            imgElement.src = URL.createObjectURL(blob);
            imgElement.style.display = 'block';

            const downloadLink = document.createElement('a');
            downloadLink.href = imgElement.src;
            downloadLink.download = relativePath.split('/').pop(); // Use the file name from the path
            downloadLink.style.display = 'none'; // Hide the download link

            imagesContainer.appendChild(imgElement);
            imagesContainer.appendChild(downloadLink);

            // Trigger the download automatically
            downloadLink.click();
        });
    } catch (error) {
        console.error('Error extracting images:', error);
        alert('Failed to extract images. Please check the console for more details.');
    }
}
