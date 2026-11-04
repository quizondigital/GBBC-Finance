(function () {
  var form = document.getElementById('gift-form');
  var catalogNode = document.getElementById('fund-catalog');
  if (!form || !catalogNode) return;

  var catalog = JSON.parse(catalogNode.textContent || '[]');
  var blocks = document.getElementById('fund-blocks');
  var fundTemplate = document.getElementById('fund-block-template');
  var subTemplate = document.getElementById('subfund-row-template');
  var totalInput = document.getElementById('gift-total');
  var methodInput = document.getElementById('gift-method');
  var referenceLabel = document.getElementById('reference-label');

  function cents(value) {
    if (value === null || value === undefined || String(value).trim() === '') return null;
    var number = Number(value);
    if (!isFinite(number)) return null;
    return Math.round(number * 100);
  }

  function money(centsValue) {
    return '$' + (centsValue / 100).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function fundById(id) {
    return catalog.filter(function (fund) { return String(fund.id) === String(id); })[0] || null;
  }

  function fillSubfundSelect(select, fundId, selected) {
    var fund = fundById(fundId);
    var current = selected === undefined ? select.value : selected;
    select.innerHTML = '';
    var placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = fund && fund.subfunds.length ? 'Select sub-fund' : 'No sub-funds yet';
    select.appendChild(placeholder);
    (fund ? fund.subfunds : []).forEach(function (sub) {
      var option = document.createElement('option');
      option.value = String(sub.id);
      option.textContent = sub.name;
      if (String(sub.id) === String(current)) option.selected = true;
      select.appendChild(option);
    });
  }

  function usedFundIds(exceptSelect) {
    var used = {};
    blocks.querySelectorAll('.fund-select').forEach(function (select) {
      if (select !== exceptSelect && select.value) used[select.value] = true;
    });
    return used;
  }

  function refreshFundChoices() {
    blocks.querySelectorAll('.fund-select').forEach(function (select) {
      var used = usedFundIds(select);
      Array.prototype.forEach.call(select.options, function (option) {
        option.disabled = !!used[option.value];
      });
    });
  }

  function bindSubfundRow(row) {
    row.querySelector('.remove-subfund').addEventListener('click', function () {
      row.remove();
      updateStatus();
    });
    row.querySelector('.subfund-amount').addEventListener('input', updateStatus);
    row.querySelector('.subfund-select').addEventListener('change', updateStatus);
  }

  function bindFundBlock(block) {
    var select = block.querySelector('.fund-select');
    select.addEventListener('change', function () {
      block.querySelectorAll('.subfund-select').forEach(function (subSelect) {
        fillSubfundSelect(subSelect, select.value, '');
      });
      refreshFundChoices();
      updateStatus();
    });
    block.querySelector('.fund-amount').addEventListener('input', updateStatus);
    block.querySelector('.remove-fund').addEventListener('click', function () {
      if (blocks.querySelectorAll('[data-fund-block]').length < 2) return;
      block.remove();
      refreshFundChoices();
      updateStatus();
    });
    block.querySelector('.add-subfund').addEventListener('click', function () {
      var row = subTemplate.content.firstElementChild.cloneNode(true);
      fillSubfundSelect(row.querySelector('.subfund-select'), select.value, '');
      block.querySelector('.subfund-list').appendChild(row);
      bindSubfundRow(row);
      updateStatus();
    });
    block.querySelectorAll('[data-subfund-row]').forEach(bindSubfundRow);
  }

  function updateReferenceLabel() {
    referenceLabel.textContent = methodInput.value === 'Check' ? 'Check number' : 'Reference';
  }

  function updateStatus() {
    var total = cents(totalInput.value);
    var allocated = 0;
    var hasAmount = false;
    var fundCount = 0;
    var subCount = 0;
    var problem = '';
    blocks.querySelectorAll('[data-fund-block]').forEach(function (block) {
      fundCount += 1;
      var fundCents = cents(block.querySelector('.fund-amount').value);
      var fundName = block.querySelector('.fund-select').selectedOptions[0].textContent;
      var note = block.querySelector('.fund-balance');
      var subSum = 0;
      var subRows = block.querySelectorAll('[data-subfund-row]');
      var subProblem = '';
      subRows.forEach(function (row) {
        subCount += 1;
        var line = cents(row.querySelector('.subfund-amount').value);
        if (!row.querySelector('.subfund-select').value) subProblem = 'Choose a sub-fund for ' + fundName + '.';
        if (line === null) subProblem = subProblem || 'Enter an amount for each sub-fund in ' + fundName + '.';
        else subSum += line;
      });
      if (fundCents !== null) {
        hasAmount = true;
        allocated += fundCents;
      }
      if (subRows.length && fundCents !== null && subSum !== fundCents) {
        subProblem = 'Sub-funds for ' + fundName + ' must add up to ' + money(fundCents) + '.';
      }
      if (!problem && fundCents === null) problem = 'Enter an amount for ' + fundName + '.';
      if (!problem && subProblem) problem = subProblem;
      if (subRows.length && fundCents !== null) {
        note.hidden = false;
        note.textContent = subSum === fundCents
          ? 'Sub-funds match ' + money(fundCents) + '.'
          : 'Sub-funds ' + money(subSum) + ' of ' + money(fundCents) + '.';
      } else {
        note.hidden = true;
        note.textContent = '';
      }
    });
    var count = document.getElementById('split-count');
    var fundLabel = fundCount + (fundCount === 1 ? ' fund' : ' funds');
    var subLabel = subCount + (subCount === 1 ? ' sub-fund' : ' sub-funds');
    count.textContent = subCount ? fundLabel + ' · ' + subLabel : fundLabel;
    var bar = document.getElementById('allocation-bar');
    var text = document.getElementById('allocation-text');
    var flag = document.getElementById('allocation-flag');
    var totalText = total === null ? '$0.00' : money(total);
    text.textContent = 'Allocated ' + (hasAmount ? money(allocated) : '$0.00') + ' / ' + totalText;
    bar.classList.remove('balanced', 'short', 'over');
    blocks.querySelectorAll('.remove-fund').forEach(function (button) {
      button.disabled = fundCount < 2;
    });
    if (problem) {
      bar.classList.add('short');
      flag.textContent = problem;
      return false;
    }
    if (total === null || !hasAmount) {
      bar.classList.add('short');
      flag.textContent = 'Enter fund amounts';
      return false;
    }
    if (allocated === total) {
      bar.classList.add('balanced');
      flag.textContent = 'Fully allocated';
      return true;
    }
    if (allocated < total) {
      bar.classList.add('short');
      flag.textContent = money(total - allocated) + ' left to allocate';
      return false;
    }
    bar.classList.add('over');
    flag.textContent = money(allocated - total) + ' over the total';
    return false;
  }

  function serialize() {
    var payload = [];
    blocks.querySelectorAll('[data-fund-block]').forEach(function (block) {
      var subfunds = [];
      block.querySelectorAll('[data-subfund-row]').forEach(function (row) {
        subfunds.push({
          subfund_id: row.querySelector('.subfund-select').value,
          amount: row.querySelector('.subfund-amount').value,
        });
      });
      payload.push({
        fund_id: block.querySelector('.fund-select').value,
        amount: block.querySelector('.fund-amount').value,
        subfunds: subfunds,
      });
    });
    document.getElementById('allocations_json').value = JSON.stringify(payload);
  }

  document.getElementById('add-fund').addEventListener('click', function () {
    var block = fundTemplate.content.firstElementChild.cloneNode(true);
    var select = block.querySelector('.fund-select');
    var used = usedFundIds(null);
    var available = Array.prototype.filter.call(select.options, function (option) { return !used[option.value]; })[0];
    if (!available) return;
    select.value = available.value;
    blocks.appendChild(block);
    bindFundBlock(block);
    refreshFundChoices();
    updateStatus();
  });

  methodInput.addEventListener('change', updateReferenceLabel);
  totalInput.addEventListener('input', updateStatus);
  blocks.querySelectorAll('[data-fund-block]').forEach(bindFundBlock);
  form.addEventListener('submit', function (event) {
    serialize();
    if (!updateStatus()) {
      event.preventDefault();
      document.getElementById('allocation-bar').scrollIntoView({ block: 'nearest' });
    }
  });
  updateReferenceLabel();
  refreshFundChoices();
  updateStatus();
})();
