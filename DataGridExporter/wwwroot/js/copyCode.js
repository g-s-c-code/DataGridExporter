function copyCode() {
    const codeElement = document.getElementById('code');
    if (codeElement) {
        navigator.clipboard.writeText(codeElement.innerText)
            .then(() => {
                alert('Code copied to clipboard!');
            })
            .catch(err => {
                console.error('Failed to copy text: ', err);
            });
    }
}