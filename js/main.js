// Hope Restoration HRF - Main JS

document.addEventListener('DOMContentLoaded', () => {
  // Mobile menu toggle
  const toggle = document.querySelector('.menu-toggle');
  const navLinks = document.querySelector('.nav-links');

  if (toggle && navLinks) {
    toggle.addEventListener('click', () => {
      navLinks.classList.toggle('open');
      toggle.setAttribute('aria-expanded', navLinks.classList.contains('open'));
    });

    // Close menu when a link is clicked
    navLinks.querySelectorAll('a').forEach(link => {
      link.addEventListener('click', () => {
        navLinks.classList.remove('open');
        toggle.setAttribute('aria-expanded', 'false');
      });
    });
  }

  // Contact Form & Direct Email Handling
  const form = document.getElementById('contact-form');
  const mailtoBtn = document.getElementById('mailto-btn');
  const statusEl = document.getElementById('form-status');
  const submitBtn = document.getElementById('submit-btn');
  const submitBtnText = document.getElementById('submit-btn-text');

  const subjectMap = {
    'support': 'Request for Support / Care',
    'partnership': 'Partnership / Collaboration',
    'volunteer': 'Volunteering',
    'general': 'General Enquiry',
    'other': 'Other Enquiry'
  };

  function showStatus(type, messageHtml) {
    if (!statusEl) return;
    statusEl.className = 'form-status ' + type;
    statusEl.innerHTML = messageHtml;
    statusEl.style.display = 'block';
    statusEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  function clearStatus() {
    if (!statusEl) return;
    statusEl.style.display = 'none';
    statusEl.innerHTML = '';
  }

  function openEmailClient(targetForm) {
    const name = targetForm ? (targetForm.querySelector('#name')?.value.trim() || '') : '';
    const email = targetForm ? (targetForm.querySelector('#email')?.value.trim() || '') : '';
    const phone = targetForm ? (targetForm.querySelector('#phone')?.value.trim() || '') : '';
    const subjectVal = targetForm ? (targetForm.querySelector('#subject')?.value || '') : '';
    const subjectLabel = subjectMap[subjectVal] || subjectVal || 'General Enquiry';
    const message = targetForm ? (targetForm.querySelector('#message')?.value.trim() || '') : '';

    const recipient = 'elijah.adebayo@hoperestorationhrf.org';
    const emailSubject = `[Hope Restoration HRF] ${subjectLabel}${name ? ' - ' + name : ''}`;

    let body = `Dear Hope Restoration and Health Relief Foundation,\n\n`;
    if (message) {
      body += `${message}\n\n`;
    } else {
      body += `I am contacting you regarding: ${subjectLabel}.\n\n`;
    }
    body += `------------------------------------\n`;
    body += `Sender Information:\n`;
    body += `Name: ${name || 'Not specified'}\n`;
    body += `Email: ${email || 'Not specified'}\n`;
    if (phone) body += `Phone: ${phone}\n`;
    body += `Subject: ${subjectLabel}\n`;
    body += `------------------------------------\n`;
    body += `Sent via Hope Restoration and Health Relief Foundation website`;

    showStatus('loading', 'Opening your email application addressed to <strong>' + recipient + '</strong>...');
    
    // Trigger mailto link
    const mailtoUrl = `mailto:${recipient}?subject=${encodeURIComponent(emailSubject)}&body=${encodeURIComponent(body)}`;
    window.location.href = mailtoUrl;

    setTimeout(() => {
      showStatus('success', 'If your email app opened, please review and hit Send. You can also email directly to <strong>' + recipient + '</strong>.');
    }, 2000);
  }

  // Handle dedicated "Open in Email App" button
  if (mailtoBtn && form) {
    mailtoBtn.addEventListener('click', () => {
      openEmailClient(form);
    });
  }

  // Handle Form Submission via FormSubmit AJAX endpoint
  if (form) {
    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      clearStatus();

      const nameInput = form.querySelector('#name');
      const emailInput = form.querySelector('#email');
      const phoneInput = form.querySelector('#phone');
      const subjectInput = form.querySelector('#subject');
      const messageInput = form.querySelector('#message');

      const name = nameInput.value.trim();
      const email = emailInput.value.trim();
      const phone = phoneInput ? phoneInput.value.trim() : '';
      const subjectVal = subjectInput.value;
      const subjectLabel = subjectMap[subjectVal] || subjectVal || 'General Enquiry';
      const message = messageInput.value.trim();

      // Client-side validations
      if (!name) {
        showStatus('error', 'Please enter your full name.');
        nameInput.focus();
        return;
      }

      const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (!email || !emailRegex.test(email)) {
        showStatus('error', 'Please enter a valid email address so we can reply to you.');
        emailInput.focus();
        return;
      }

      if (!subjectVal) {
        showStatus('error', 'Please select a subject for your enquiry.');
        subjectInput.focus();
        return;
      }

      if (!message) {
        showStatus('error', 'Please enter your message.');
        messageInput.focus();
        return;
      }

      // Set sending UI state
      const originalBtnText = submitBtnText ? submitBtnText.textContent : 'Send Message Directly';
      if (submitBtn) submitBtn.disabled = true;
      if (submitBtnText) submitBtnText.textContent = 'Sending Message...';
      showStatus('loading', 'Sending your message to <strong>elijah.adebayo@hoperestorationhrf.org</strong>...');

      const payload = {
        name: name,
        email: email,
        phone: phone || 'Not provided',
        subject: subjectLabel,
        message: message,
        _subject: `New Website Enquiry: ${subjectLabel} from ${name}`,
        _replyto: email,
        _template: 'table',
        _captcha: 'false'
      };

      try {
        const response = await fetch('https://formsubmit.co/ajax/elijah.adebayo@hoperestorationhrf.org', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Accept': 'application/json'
          },
          body: JSON.stringify(payload)
        });

        const data = await response.json().catch(() => ({}));

        if (response.ok && (data.success === 'true' || data.success === true || response.status === 200)) {
          showStatus(
            'success',
            '✓ Thank you! Your message has been sent successfully to <strong>elijah.adebayo@hoperestorationhrf.org</strong>. We will respond within 1 to 2 working days.'
          );
          form.reset();
          if (submitBtnText) submitBtnText.textContent = 'Message Sent ✓';
          setTimeout(() => {
            if (submitBtnText) submitBtnText.textContent = originalBtnText;
            if (submitBtn) submitBtn.disabled = false;
          }, 4000);
        } else {
          throw new Error(data.message || 'Server returned an error');
        }
      } catch (err) {
        console.warn('Direct web send issue, opening email fallback:', err);
        showStatus(
          'error',
          'Network submission encountered a delay. <a href="#" id="error-mailto-link" style="color: #991b1b; text-decoration: underline; font-weight: 600;">Click here to send directly via your email app (Gmail / Outlook)</a>, or email <strong>elijah.adebayo@hoperestorationhrf.org</strong>.'
        );
        const errorMailtoLink = document.getElementById('error-mailto-link');
        if (errorMailtoLink) {
          errorMailtoLink.addEventListener('click', (ev) => {
            ev.preventDefault();
            openEmailClient(form);
          });
        }
        if (submitBtnText) submitBtnText.textContent = originalBtnText;
        if (submitBtn) submitBtn.disabled = false;
      }
    });
  }

  // Set active nav link based on current page
  const currentPage = window.location.pathname.split('/').pop() || 'index.html';
  document.querySelectorAll('.nav-links a').forEach(link => {
    const href = link.getAttribute('href');
    if (href === currentPage || (currentPage === '' && href === 'index.html')) {
      link.classList.add('active');
    }
  });
});
