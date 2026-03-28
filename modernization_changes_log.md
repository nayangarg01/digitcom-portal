# Digitcom Website Modernization: Summary of Changes

This document tracks all modifications made during the modernization phase. You can provide this to a developer or AI agent during the final Render migration to ensure all improvements are preserved.

## 1. Backend & Modern Form Handling
*   **Database Schema**: Added the `Lead` model to `prisma/schema.prisma` to store contact form submissions in PostgreSQL.
*   **API Routes**: 
    *   Created `src/controllers/contactController.ts` to handle form validation and database storage.
    *   Created `src/routes/contactRoutes.ts` with the `/api/contact/submit` endpoint.
    *   Registered routes in `src/index.ts`.
*   **AJAX Submission**: 
    *   Created `scripts/contact.js` to handle form submissions without page reloads.
    *   Updated `contact.html` to remove legacy `action="contact.php"` and link the new JavaScript handler.

## 2. UI/UX Refinements (Global)
*   **Header Padding**: Reduced `.page-header` vertical padding from `160px` to `100px` across all sub-pages (`gallery.html`, `clients.html`, `company-profile.html`, etc.) to improve content visibility.
*   **Typography**: Increased the font size of the hero badge text on `index.html` from `0.9rem` to `1.1rem` for better readability.
*   **Form Cleanup**: Removed all placeholder text from the `contact.html` form inputs for a cleaner, unified Look.

## 3. Content Corrections
*   **Gallery Organization**: Relocated 10 images from the miscategorized "Mobile Networks" middle-section into the "Solar" category in `gallery.html`.
*   **Client Logos**: Updated image paths for **Reliance Jio** and **92.7 Big FM** in `clients.html` to use lowercase filenames, ensuring they load correctly on all Linux-based servers (like GoDaddy/cPanel).

## 4. Professionalism & Performance
*   **Portal Login**: Updated `portal/js/auth.js` to remove "Render wake-up" warnings from the login button, providing a smoother experience for users as the backend transitions to a paid tier.

---
*Last Updated: March 28, 2026*
