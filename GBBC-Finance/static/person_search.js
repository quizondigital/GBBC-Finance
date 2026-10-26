(function () {
  var root = document.querySelector('.person-search');
  if (!root) return;

  var searchUrl = root.getAttribute('data-search-url');
  var input = document.getElementById('person_search');
  var hidden = document.getElementById('person_id');
  var list = document.getElementById('person_results');
  var form = document.getElementById('gift-form');
  var timer = null;
  var activeIndex = -1;
  var results = [];
  var selectedLabel = (input && input.value) || '';

  function setExpanded(open) {
    input.setAttribute('aria-expanded', open ? 'true' : 'false');
    list.hidden = !open;
  }

  function clearResults() {
    results = [];
    activeIndex = -1;
    list.innerHTML = '';
    setExpanded(false);
  }

  function highlight(name, query) {
    var safe = name.replace(/[&<>"']/g, function (ch) {
      return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[ch];
    });
    var tokens = query.trim().split(/\s+/).filter(Boolean);
    if (!tokens.length) return safe;
    var pattern = new RegExp('(' + tokens.map(function (t) {
      return t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }).join('|') + ')', 'ig');
    return safe.replace(pattern, '<mark>$1</mark>');
  }

  function render() {
    list.innerHTML = '';
    if (!results.length) {
      var empty = document.createElement('li');
      empty.className = 'person-search-empty';
      empty.textContent = 'No matching people';
      list.appendChild(empty);
      setExpanded(true);
      return;
    }
    results.forEach(function (person, index) {
      var item = document.createElement('li');
      item.setAttribute('role', 'option');
      item.id = 'person-option-' + person.id;
      item.dataset.id = String(person.id);
      item.dataset.name = person.name;
      if (index === activeIndex) item.classList.add('is-active');
      item.innerHTML =
        '<span class="person-search-name">' + highlight(person.name, input.value) + '</span>' +
        '<small>' + (person.member_type || '') + '</small>';
      item.addEventListener('mousedown', function (event) {
        event.preventDefault();
        selectPerson(person);
      });
      list.appendChild(item);
    });
    setExpanded(true);
  }

  function selectPerson(person) {
    hidden.value = String(person.id);
    selectedLabel = person.name;
    input.value = person.name;
    clearResults();
  }

  function fetchMatches(query) {
    if (query.length < 2) {
      clearResults();
      return;
    }
    fetch(searchUrl + '?q=' + encodeURIComponent(query), {
      headers: { Accept: 'application/json' },
      credentials: 'same-origin',
    })
      .then(function (response) {
        if (!response.ok) throw new Error('Search failed');
        return response.json();
      })
      .then(function (data) {
        if (input.value.trim() !== query) return;
        results = Array.isArray(data) ? data : [];
        activeIndex = results.length ? 0 : -1;
        render();
      })
      .catch(function () {
        clearResults();
      });
  }

  input.addEventListener('input', function () {
    var query = input.value.trim();
    if (selectedLabel && input.value !== selectedLabel) {
      hidden.value = '';
      selectedLabel = '';
    }
    window.clearTimeout(timer);
    if (query.length < 2) {
      clearResults();
      return;
    }
    timer = window.setTimeout(function () {
      fetchMatches(query);
    }, 180);
  });

  input.addEventListener('keydown', function (event) {
    if (list.hidden || !results.length) {
      if (event.key === 'Escape') clearResults();
      return;
    }
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      activeIndex = (activeIndex + 1) % results.length;
      render();
    } else if (event.key === 'ArrowUp') {
      event.preventDefault();
      activeIndex = (activeIndex - 1 + results.length) % results.length;
      render();
    } else if (event.key === 'Enter' && activeIndex >= 0) {
      event.preventDefault();
      selectPerson(results[activeIndex]);
    } else if (event.key === 'Escape') {
      clearResults();
    }
  });

  input.addEventListener('blur', function () {
    window.setTimeout(clearResults, 120);
  });

  if (form) {
    form.addEventListener('submit', function (event) {
      if (!hidden.value) {
        event.preventDefault();
        input.focus();
        window.alert('Choose a contributor from the search results.');
      }
    });
  }
})();
