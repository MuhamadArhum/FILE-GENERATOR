// Show loading state when form is being submitted
document.querySelector('form').addEventListener('submit', function(e) {
    e.preventDefault();  // Prevent form submission to show spinner

    let button = document.querySelector('button');
    let spinner = document.getElementById('loading-spinner');
    let form = document.getElementById('upload-form');
    
    // Show spinner and disable the button
    spinner.style.display = 'block';
    button.disabled = true;
    button.innerHTML = 'Uploading...';

    // Simulate file upload (for testing)
    setTimeout(function() {
        form.submit(); // Proceed with actual submission after spinner is shown
    }, 3000); // Simulate a 3-second upload delay
});
