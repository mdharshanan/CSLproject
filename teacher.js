document.addEventListener('DOMContentLoaded', function () {
  const savedCoursesKey = 'teacherCreatedCourses';
  const savedResultsKey = 'teacherPublishedResults';
  const savedTeachersKey = 'teacherProfiles';

  function setMessage(selector, text, isError) {
    const element = document.querySelector(selector);
    if (!element) {
      return;
    }

    element.textContent = text;
    element.classList.toggle('error', Boolean(isError));
  }

  function readList(key) {
    try {
      return JSON.parse(localStorage.getItem(key)) || [];
    } catch (error) {
      return [];
    }
  }

  function saveList(key, value) {
    localStorage.setItem(key, JSON.stringify(value));
  }

  function renderList(selector, items, formatter) {
    const list = document.querySelector(selector);
    if (!list) {
      return;
    }

    list.innerHTML = '';
    items.forEach(function (item) {
      const li = document.createElement('li');
      li.textContent = formatter(item);
      list.appendChild(li);
    });
  }

  function renderSavedData() {
    renderList('[data-teacher-course-list]', readList(savedCoursesKey), function (course) {
      return `${course.title} - ${course.category} - starts ${course.startDate}`;
    });

    renderList('[data-teacher-result-list]', readList(savedResultsKey), function (result) {
      return `${result.student}: ${result.marks} marks`;
    });

    renderTeacherProfiles();
  }

  function addCourseToPage(course) {
    const courses = readList(savedCoursesKey);
    courses.unshift(course);
    saveList(savedCoursesKey, courses);
    renderSavedData();
  }

  function renderTeacherProfiles() {
    const table = document.getElementById('teacher-profiles-table');
    if (!table) {
      return;
    }

    table.querySelectorAll('[data-added-teacher]').forEach(function (row) {
      row.remove();
    });

    readList(savedTeachersKey).forEach(function (teacher) {
      const row = document.createElement('tr');
      row.setAttribute('data-added-teacher', 'true');
      [teacher.name, 'Registered Teacher', teacher.department, 'Pending', teacher.email].forEach(function (value) {
        const cell = document.createElement('td');
        cell.textContent = value;
        row.appendChild(cell);
      });
      table.appendChild(row);
    });
  }

  const loginForm = document.getElementById('teacher-login-form');
  if (loginForm) {
    loginForm.addEventListener('submit', function (event) {
      event.preventDefault();
      const email = document.getElementById('teacher-login-email').value.trim();
      const password = document.getElementById('teacher-login-password').value.trim();

      if (!email || !password) {
        setMessage('[data-teacher-login-message]', 'Please enter teacher email and password.', true);
        return;
      }

      setMessage('[data-teacher-login-message]', 'Teacher dashboard unlocked for this session.', false);
      loginForm.reset();
    });
  }

  const registerForm = document.getElementById('teacher-register-form');
  if (registerForm) {
    registerForm.addEventListener('submit', function (event) {
      event.preventDefault();

      const teacher = {
        name: document.getElementById('teacher-register-name').value.trim(),
        email: document.getElementById('teacher-register-email').value.trim().toLowerCase(),
        department: document.getElementById('teacher-register-department').value.trim(),
      };
      const password = document.getElementById('teacher-register-password').value.trim();

      if (!teacher.name || !teacher.email || !teacher.department || !password) {
        setMessage('[data-teacher-register-message]', 'Please fill all teacher registration fields.', true);
        return;
      }

      const teachers = readList(savedTeachersKey);
      teachers.push(teacher);
      saveList(savedTeachersKey, teachers);
      renderSavedData();
      setMessage('[data-teacher-register-message]', 'Teacher registered successfully.', false);
      registerForm.reset();
    });
  }

  const courseForm = document.getElementById('teacher-course-form');
  if (courseForm) {
    courseForm.addEventListener('submit', async function (event) {
      event.preventDefault();

      const course = {
        title: document.getElementById('teacher-course-title').value.trim(),
        category: document.getElementById('teacher-course-category').value.trim(),
        teacher: document.getElementById('teacher-course-teacher').value.trim(),
        duration: document.getElementById('teacher-course-duration').value.trim(),
        startDate: document.getElementById('teacher-course-start').value,
        description: document.getElementById('teacher-course-description').value.trim(),
      };

      if (!course.title || !course.category || !course.teacher || !course.duration || !course.startDate) {
        setMessage('[data-teacher-course-message]', 'Please fill all required course fields.', true);
        return;
      }

      try {
        const response = await fetch('/api/teacher/courses', {
          method: 'POST',
          credentials: 'same-origin',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(course),
        });
        const result = await response.json().catch(function () {
          return {};
        });

        if (!response.ok) {
          addCourseToPage(course);
          setMessage('[data-teacher-course-message]', result.error || 'Course added on this page. Restart backend to save it in database.', false);
          courseForm.reset();
          return;
        }

        addCourseToPage(result.course || course);
        setMessage('[data-teacher-course-message]', 'Course created successfully.', false);
        courseForm.reset();
      } catch (error) {
        console.error(error);
        addCourseToPage(course);
        setMessage('[data-teacher-course-message]', 'Course added on this page. Start the backend to also save it in database.', false);
        courseForm.reset();
      }
    });
  }

  const materialForm = document.getElementById('teacher-material-form');
  if (materialForm) {
    materialForm.addEventListener('submit', function (event) {
      event.preventDefault();
      const file = document.getElementById('teacher-material-file').files[0];
      if (!file) {
        setMessage('[data-teacher-material-message]', 'Please choose a notes file.', true);
        return;
      }

      setMessage('[data-teacher-material-message]', `Notes selected: ${file.name}`, false);
      materialForm.reset();
    });
  }

  const videoForm = document.getElementById('teacher-video-form');
  if (videoForm) {
    videoForm.addEventListener('submit', function (event) {
      event.preventDefault();
      const file = document.getElementById('teacher-video-file').files[0];
      if (!file) {
        setMessage('[data-teacher-video-message]', 'Please choose a video file.', true);
        return;
      }

      setMessage('[data-teacher-video-message]', `Video selected: ${file.name}`, false);
      videoForm.reset();
    });
  }

  document.querySelectorAll('.review-assignment').forEach(function (button) {
    button.addEventListener('click', function () {
      const row = button.closest('tr');
      const statusCell = row ? row.children[2] : null;
      if (statusCell) {
        statusCell.textContent = 'Reviewed';
      }
      button.textContent = 'Reviewed';
      button.disabled = true;
      setMessage('[data-teacher-assignment-message]', 'Assignment marked as reviewed.', false);
    });
  });

  const resultForm = document.getElementById('teacher-result-form');
  if (resultForm) {
    resultForm.addEventListener('submit', function (event) {
      event.preventDefault();

      const result = {
        student: document.getElementById('teacher-result-student').value.trim(),
        marks: document.getElementById('teacher-result-marks').value.trim(),
      };

      if (!result.student || !result.marks) {
        setMessage('[data-teacher-result-message]', 'Please enter student name and marks.', true);
        return;
      }

      const results = readList(savedResultsKey);
      results.unshift(result);
      saveList(savedResultsKey, results);
      renderSavedData();
      setMessage('[data-teacher-result-message]', 'Result published successfully.', false);
      resultForm.reset();
    });
  }

  renderSavedData();
});
