document.addEventListener('DOMContentLoaded', async () => {
  const jobsCount = document.getElementById('jobsCount');
  const customerCount = document.getElementById('customerCount');
  const fieldCount = document.getElementById('fieldCount');
  const strawCount = document.getElementById('strawCount');

  async function loadCounts() {
    try {
      const jobsRes = await fetch('/api/jobs', { cache: 'no-store' });
      const jobsData = jobsRes.ok ? await jobsRes.json() : { jobs: [] };
      const jobs = Array.isArray(jobsData.jobs) ? jobsData.jobs : [];

      const customerRes = await fetch('/api/customers', { cache: 'no-store' });
      const customerData = customerRes.ok ? await customerRes.json() : { customers: [] };
      const customers = Array.isArray(customerData.customers) ? customerData.customers : [];

      const fieldMapRes = await fetch('/api/field-map', { cache: 'no-store' });
      const fieldMapData = fieldMapRes.ok ? await fieldMapRes.json() : { field_map: {} };
      const fieldMap = fieldMapData.field_map || {};
      let totalFields = 0;
      Object.values(fieldMap).forEach((customerValue) => {
        if (!customerValue || typeof customerValue !== 'object') return;
        Object.values(customerValue).forEach((farmValue) => {
          if (Array.isArray(farmValue)) totalFields += farmValue.length;
        });
      });

      if (jobsCount) jobsCount.textContent = String(jobs.length);
      if (customerCount) customerCount.textContent = String(customers.length);
      if (fieldCount) fieldCount.textContent = String(totalFields);
      if (strawCount) strawCount.textContent = 'On';
    } catch (error) {
      console.error('Unable to load dashboard metrics', error);
      if (jobsCount) jobsCount.textContent = '—';
      if (customerCount) customerCount.textContent = '—';
      if (fieldCount) fieldCount.textContent = '—';
      if (strawCount) strawCount.textContent = '—';
    }
  }

  await loadCounts();
});
