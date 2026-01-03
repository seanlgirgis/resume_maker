document.addEventListener('DOMContentLoaded', () => {
    // Auto-highlight active nav link
    const currentPath = window.location.pathname;
    const navLinks = document.querySelectorAll('.nav-menu a');
    
    navLinks.forEach(link => {
        if (currentPath.includes(link.getAttribute('href'))) {
            link.classList.add('active');
        }
    });

    console.log("Sean Luka Girgis Portfolio Engine Initialized.");
});