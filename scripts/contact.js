const API_URL = 'https://digitcom-backend.onrender.com/api';

document.addEventListener('DOMContentLoaded', () => {
    const contactForm = document.querySelector('.contact-form form');
    if (!contactForm) return;

    // Create message element
    const statusMsg = document.createElement('div');
    statusMsg.style.marginTop = '20px';
    statusMsg.style.padding = '15px';
    statusMsg.style.borderRadius = '8px';
    statusMsg.style.display = 'none';
    contactForm.appendChild(statusMsg);

    contactForm.addEventListener('submit', async (e) => {
        e.preventDefault();

        const submitBtn = contactForm.querySelector('button[type="submit"]');
        const originalText = submitBtn.textContent;

        // Collect data
        const formData = {
            name: document.getElementById('name').value,
            email: document.getElementById('email').value,
            company: document.getElementById('company').value || '',
            message: document.getElementById('message').value
        };

        // UI Loading state
        submitBtn.disabled = true;
        submitBtn.textContent = 'Sending...';
        statusMsg.style.display = 'none';

        try {
            const response = await fetch(`${API_URL}/contact/submit`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(formData)
            });

            const data = await response.json();

            if (response.ok) {
                // Success
                statusMsg.textContent = data.message || 'Thank you! Your message has been sent.';
                statusMsg.style.backgroundColor = '#dcfce7';
                statusMsg.style.color = '#166534';
                statusMsg.style.border = '1px solid #bbf7d0';
                statusMsg.style.display = 'block';
                contactForm.reset();
            } else {
                // Error from server
                throw new Error(data.error || 'Failed to send message.');
            }
        } catch (error) {
            console.error('Contact error:', error);
            statusMsg.textContent = error.message.includes('fetch') 
                ? 'Server connection failed. The backend might be starting up (20-40s).' 
                : error.message;
            statusMsg.style.backgroundColor = '#fee2e2';
            statusMsg.style.color = '#991b1b';
            statusMsg.style.border = '1px solid #fecaca';
            statusMsg.style.display = 'block';
        } finally {
            submitBtn.disabled = false;
            submitBtn.textContent = originalText;
        }
    });
});
