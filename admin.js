document.addEventListener('DOMContentLoaded', function () {
  const adminForm = document.querySelector('.admin-card');
  const messageContainer = document.createElement('p');
  messageContainer.style.marginTop = '12px';
  messageContainer.style.color = '#111';
  adminForm.appendChild(messageContainer);

  if (!adminForm) {
    return;
  }

  adminForm.addEventListener('submit', async function (event) {
    event.preventDefault();

    const data = {
      title: document.querySelector('#course-title')?.value.trim(),
      category: document.querySelector('#course-category')?.value.trim(),
      teacher: document.querySelector('#course-teacher')?.value.trim(),
      duration: document.querySelector('#course-duration')?.value.trim(),
      startDate: document.querySelector('#course-start')?.value,
      certificateEnabled: document.querySelector('#certificate-enabled')?.checked,
      certificateRequirement: document.querySelector('#certificate-requirement')?.value,
      description: document.querySelector('#course-description')?.value.trim(),
    };

    try {
      const response = await fetch('/api/courses', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(data),
      });

      const result = await response.json();

      if (!response.ok) {
        messageContainer.textContent = result.error || 'Unable to create course.';
        messageContainer.style.color = '#d32f2f';
        return;
      }

      messageContainer.textContent = 'Course created successfully. A notification email was sent.';
      messageContainer.style.color = '#1e7e34';
      adminForm.reset();
    } catch (error) {
      messageContainer.textContent = 'Server error creating the course.';
      messageContainer.style.color = '#d32f2f';
      console.error(error);
    }
  });
});
