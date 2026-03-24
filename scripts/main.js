/**
 * Main JavaScript for Digitcom India Technologies
 * Handles Mobile Menu, Sticky Navbar, and Dropdown behavior
 */

document.addEventListener('DOMContentLoaded', () => {
    const mobileMenuBtn = document.querySelector('.mobile-menu-btn');
    const navLinks = document.querySelector('.nav-links');
    const navbar = document.querySelector('.navbar');
    const dropdowns = document.querySelectorAll('.dropdown');

    // --- Mobile Menu Toggle ---
    if (mobileMenuBtn && navLinks) {
        mobileMenuBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            navLinks.classList.toggle('active');
            mobileMenuBtn.classList.toggle('active');
            document.body.classList.toggle('menu-open');
            
            // Toggle hamburger icon (if you want to use text/icons instead of CSS animation)
            if (mobileMenuBtn.classList.contains('active')) {
                mobileMenuBtn.innerHTML = '✕';
            } else {
                mobileMenuBtn.innerHTML = '☰';
            }
        });
    }

    // --- Mobile Dropdown Toggle ---
    dropdowns.forEach(dropdown => {
        const dropbtn = dropdown.querySelector('.dropbtn');
        const content = dropdown.querySelector('.dropdown-content');
        
        if (dropbtn && content) {
            dropbtn.addEventListener('click', (e) => {
                if (window.innerWidth <= 768) {
                    e.preventDefault();
                    e.stopPropagation();
                    
                    // Close other dropdowns
                    dropdowns.forEach(d => {
                        if (d !== dropdown) {
                            d.querySelector('.dropdown-content').classList.remove('active');
                        }
                    });
                    
                    content.classList.toggle('active');
                }
            });
        }
    });

    // --- Sticky Navbar Effect ---
    const handleScroll = () => {
        if (window.scrollY > 50) {
            navbar.classList.add('scrolled');
        } else {
            navbar.classList.remove('scrolled');
        }
    };

    window.addEventListener('scroll', handleScroll);
    handleScroll(); // Initial check

    // --- Close menu on link click ---
    const navLinksItems = document.querySelectorAll('.nav-links a');
    navLinksItems.forEach(item => {
        item.addEventListener('click', () => {
            if (window.innerWidth <= 768) {
                navLinks.classList.remove('active');
                mobileMenuBtn.classList.remove('active');
                mobileMenuBtn.innerHTML = '☰';
                document.body.classList.remove('menu-open');
            }
        });
    });

    // --- Close menu on outside click ---
    document.addEventListener('click', (e) => {
        if (window.innerWidth <= 768 && 
            navLinks.classList.contains('active') && 
            !navLinks.contains(e.target) && 
            !mobileMenuBtn.contains(e.target)) {
            navLinks.classList.remove('active');
            mobileMenuBtn.classList.remove('active');
            mobileMenuBtn.innerHTML = '☰';
            document.body.classList.remove('menu-open');
        }
    });
});
