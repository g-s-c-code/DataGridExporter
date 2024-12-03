function saveAsFile(filename, bytesBase64, mimeType = 'application/octet-stream') {
    try {
        if (!filename || !bytesBase64) {
            console.error('Filename and base64 data are required');
            return false;
        }

        const byteCharacters = atob(bytesBase64);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        const blob = new Blob([byteArray], { type: mimeType });

        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = filename;

        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        URL.revokeObjectURL(link.href);

        return true;
    } catch (error) {
        console.error('Error saving file:', error);
        return false;
    }
}