

    function loadSavedArray(key) {

        try {

            let value = JSON.parse(localStorage.getItem(key));

            return Array.isArray(value) ? value : [];

        } catch (error) {

            console.warn(`Invalid saved data for ${key}. Resetting this list.`, error);

            localStorage.removeItem(key);

            return [];

        }

    }



    function escapeHtml(value) {

        return String(value ?? '').replace(/[&<>"']/g, function(char) {

            return {

                '&': '&amp;',

                '<': '&lt;',

                '>': '&gt;',

                '"': '&quot;',

                "'": '&#39;'

            }[char];

        });

    }



    function encodeForClick(value) {

        return encodeURIComponent(String(value ?? ''));

    }



    function decodeFromClick(value) {

        return decodeURIComponent(value);

    }



    function createId(prefix) {

        return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;

    }



    function ensureJobIds(list) {

        let changed = false;

        list.forEach(job => {

            if(!job.id) {

                job.id = createId(job.paymentHistory ? 'pay' : 'job');

                changed = true;

            }

        });

        return changed;

    }



    function getPaymentsForJob(job) {

        if(!job || !job.id) return [];

        return jobs.filter(item => item.paymentHistory && item.relatedJobId === job.id);

    }



    function getPaidAmountForJob(job) {

        return getPaymentsForJob(job).reduce((sum, payment) => sum + (parseFloat(payment.price) || 0), 0);

    }



    function getOutstandingAmount(job) {

        if(!job || job.status !== "Unpaid") return 0;

        return Math.max(0, (parseFloat(job.price) || 0) - getPaidAmountForJob(job));

    }



    function isIncomeEntry(job) {

        if(!job) return false;

        if(job.paymentHistory) return job.status === "Paid";

        return job.status === "Paid" && getPaidAmountForJob(job) === 0;

    }



    function getIncomeAmount(job) {

        return isIncomeEntry(job) ? (parseFloat(job.price) || 0) : 0;

    }



    function getIncomeDate(job) {

        return new Date(job.payDate || job.date);

    }



    let customers = loadSavedArray('crm_customers');

    let jobs = loadSavedArray('crm_jobs');

    if(ensureJobIds(jobs)) {

        localStorage.setItem('crm_jobs', JSON.stringify(jobs));

    }

    let editIndex = null; 

    let editCustIndex = null; 

    let currentActiveModalCustomer = "";

    let currentCalculatedTotalDue = 0;



    window.onload = function() {

        if(document.getElementById('metricsYearSearch')) {

            document.getElementById('metricsYearSearch').value = new Date().getFullYear();

        }

        renderCustomers();

        renderJobs();

        updateDashboardMetrics();

        setDate();

        populatePivotDropdown();

        initializeUiEvents();

        

        document.addEventListener('click', function(event) {

            let wrapper = document.querySelector('.search-wrapper');

            let dropdown = document.getElementById("customerList");

            if (wrapper && dropdown && !wrapper.contains(event.target)) {

                dropdown.style.display = "none";

            }

        });
        document.body.addEventListener('click', function(event) {
            let target = event.target.closest('[data-target],[data-action]');
            if (!target) return;

            if (target.dataset.target) {
                event.preventDefault();
                show(target.dataset.target);
                return;
            }

            if (target.dataset.action) {
                event.preventDefault();
                let action = target.dataset.action;
                if (action && typeof window[action] === 'function') {
                    window[action]();
                }
            }
        });

    };



    function saveState(){

        localStorage.setItem('crm_customers', JSON.stringify(customers));

        localStorage.setItem('crm_jobs', JSON.stringify(jobs));

        saveStateToGoogleSheets();

    }

    function isAppsScriptContext() {
        return typeof google !== 'undefined' && google.script && google.script.run;
    }

    function loadStateFromGoogleSheets() {
        return new Promise(resolve => {
            if (!isAppsScriptContext()) {
                resolve(false);
                return;
            }

            google.script.run
                .withSuccessHandler(function(data) {
                    if (!data || !Array.isArray(data.customers) || !Array.isArray(data.jobs)) {
                        resolve(false);
                        return;
                    }

                    customers = data.customers;
                    jobs = data.jobs;

                    localStorage.setItem('crm_customers', JSON.stringify(customers));
                    localStorage.setItem('crm_jobs', JSON.stringify(jobs));
                    if (ensureJobIds(jobs)) {
                        localStorage.setItem('crm_jobs', JSON.stringify(jobs));
                    }

                    refreshAllViews();
                    resolve(true);
                })
                .withFailureHandler(function() {
                    resolve(false);
                })
                .getDataFromSheet();
        });
    }

    function saveStateToGoogleSheets() {
        if (!isAppsScriptContext()) {
            return;
        }

        google.script.run
            .withFailureHandler(function(error) {
                console.error('Google Sheets save failed:', error);
            })
            .saveDataToSheet({ customers, jobs });
    }

    async function autoSyncToGoogleSheets() {
        if (isAppsScriptContext()) {
            saveStateToGoogleSheets();
            console.log('? Auto-synced to Google Sheets via Apps Script');
            return;
        }

        const scriptURL = 'https://script.google.com/macros/s/AKfycbwLl5nVHJp_-g3nzzBWTjwg_bMdV1c-KSaIAyFGaypyB5lhx_0aYYb-KHq9iYqYL0GHgg/exec';

        try {
            let customersData = customers.map(c => [
                c.name,
                c.mobile,
                c.email || '',
                c.shop || '',
                c.address || ''
            ]);

            let jobsData = jobs.map(j => [
                j.id || '',
                j.relatedJobId || '',
                j.date || '',
                j.cname || '',
                j.description || '—',
                j.size || '—',
                j.price || 0,
                j.status || '',
                j.payDate || '',
                j.paymentHistory ? 'TRUE' : 'FALSE'
            ]);

            await fetch(scriptURL, {
                method: 'POST',
                mode: 'no-cors',
                body: JSON.stringify({ 
                    action: 'sync',
                    customers: customersData,
                    jobs: jobsData,
                    timestamp: new Date().toISOString()
                })
            });
            console.log('? Auto-synced to Google Sheets');
        } catch (error) {
            console.error('? Auto-sync error:', error);
        }
    }



    function refreshAllViews() {

        renderCustomers();

        renderJobs();

        populatePivotDropdown();

        updateDashboardMetrics();

        generateFullReport();

        generateMonthlyMatrixReport();

    }



    function show(id){

        let section = document.getElementById(id);

        let navLink = document.getElementById('nav-' + id);

        

        if(!section) return;

        

        document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));

        document.querySelectorAll('.menu a').forEach(a => a.classList.remove('active-nav'));

        

        section.classList.add('active');

        if(navLink) navLink.classList.add('active-nav');

        

        if(id === 'dashboard') updateDashboardMetrics();

        if(id === 'metrics') generateMonthlyMatrixReport();

        if(id === 'customer') renderCustomers(); 

        if(id === 'dataentry') {
 setDate();

            populatePivotDropdown();

            resetJobFilters(); 

        }

        if(id === 'report') generateFullReport();

    }



    function initializeUiEvents() {
        document.querySelectorAll('[data-target]').forEach(element => {
            element.addEventListener('click', function(event) {
                event.preventDefault();
                let target = this.dataset.target;
                if (target) show(target);
            });
        });

        if (document.getElementById('btn-clear-system')) {
            document.getElementById('btn-clear-system').addEventListener('click', clearSystemData);
        }

        if (document.getElementById('btn-submit-cust')) {
            document.getElementById('btn-submit-cust').addEventListener('click', addCustomer);
        }

        if (document.getElementById('btn-submit-job')) {
            document.getElementById('btn-submit-job').addEventListener('click', addJob);
        }

        if (document.getElementById('cname')) {
            document.getElementById('cname').addEventListener('click', toggleCustomerDropdown);
        }

        if (document.getElementById('btn-settlement')) {
            document.getElementById('btn-settlement').addEventListener('click', processSmartSettlement);
        }

        if (document.getElementById('modalPrintBtn')) {
            document.getElementById('modalPrintBtn').addEventListener('click', function() {
                window.print();
            });
        }

        document.querySelectorAll('[data-action]').forEach(element => {
            element.addEventListener('click', function(event) {
                event.preventDefault();
                let action = this.dataset.action;
                if (action && typeof window[action] === 'function') {
                    window[action]();
                }
            });
        });
    }

    function renderJobs() {

        filterJobsTable();

    }



    function generateMonthlyMatrixReport() {

        let tableBody = document.getElementById("metricsMatrixTableBody");

        if(!tableBody) return;

        tableBody.innerHTML = "";



        let yearInput = document.getElementById("metricsYearSearch")?.value;

        let selectedYear = parseInt(yearInput);



        if(isNaN(selectedYear)) {

            tableBody.innerHTML = `<tr><td colspan="14" style="text-align:center; color:var(--danger);">Please enter a valid year.</td></tr>`;

            return;

        }



        if(customers.length === 0) {

            tableBody.innerHTML = `<tr><td colspan="14" style="text-align:center; color:#999;">No clients found. Register profiles to view data.</td></tr>`;

            return;

        }



        customers.forEach(c => {

            let monthlyStats = Array.from({ length: 12 }, () => ({
                work: 0,
                paid: 0,
                unpaid: 0,
                loss: 0
            }));

            jobs.filter(j => j.cname === c.name).forEach(j => {

                if(isIncomeEntry(j)) {

                    let paidDate = getIncomeDate(j);

                    if(paidDate.getFullYear() === selectedYear) {

                        monthlyStats[paidDate.getMonth()].paid += getIncomeAmount(j);

                    }

                    return;

                }

                if(j.paymentHistory) return;

                let jobDate = new Date(j.date);

                if(jobDate.getFullYear() !== selectedYear) return;

                let monthIndex = jobDate.getMonth();

                let amount = parseFloat(j.price) || 0;

                monthlyStats[monthIndex].work += amount;

                if(j.status === "Unpaid") {

                    monthlyStats[monthIndex].unpaid += getOutstandingAmount(j);

                } else if(j.status === "Cancelled") {

                    monthlyStats[monthIndex].loss += amount;

                }

            });

            let rowTotal = monthlyStats.reduce((total, item) => ({
                work: total.work + item.work,
                paid: total.paid + item.paid,
                unpaid: total.unpaid + item.unpaid,
                loss: total.loss + item.loss
            }), { work: 0, paid: 0, unpaid: 0, loss: 0 });



            let trHtml = `<tr><td><strong>${escapeHtml(c.name)}</strong></td>`;

            for(let m = 0; m < 12; m++) {

                trHtml += `<td>${formatMetricLines(monthlyStats[m])}</td>`;

            }

            trHtml += `<td style="background-color: #f8fafc;">${formatMetricLines(rowTotal)}</td></tr>`;

            

            tableBody.innerHTML += trHtml;

        });
    }



    function formatMetricLines(item) {

        return `
            <div class="metric-lines">
                <span class="work"><b>W</b> ?${item.work.toFixed(2)}</span>
                <span class="paid"><b>P</b> ?${item.paid.toFixed(2)}</span>
                <span class="unpaid"><b>U</b> ?${item.unpaid.toFixed(2)}</span>
                <span class="loss-text"><b>L</b> ?${item.loss.toFixed(2)}</span>
            </div>
        `;

    }



    function updateDashboardMetrics(){

        let totalPaid = 0, totalDue = 0, totalLoss = 0, totalAdvance = 0, pendingJobsCount = 0;

        

        jobs.forEach(j => {

            let val = parseFloat(j.price) || 0;

            if(isIncomeEntry(j)) {

                totalPaid += getIncomeAmount(j);

                if(j.paymentHistory && !j.relatedJobId) {

                    totalAdvance += getIncomeAmount(j);

                }

            } else if(j.status === "Unpaid") {

                let dueAmount = getOutstandingAmount(j);

                totalDue += dueAmount;

                if(dueAmount > 0) pendingJobsCount++;

            } else if(j.status === "Cancelled") {

                totalLoss += val;

            }

        });



        if(document.getElementById('db-cust-count')) document.getElementById('db-cust-count').innerText = customers.length;

        let realJobCount = jobs.filter(j => !j.paymentHistory).length;

        if(document.getElementById('db-job-count')) document.getElementById('db-job-count').innerText = realJobCount;

        if(document.getElementById('db-pending-jobs')) document.getElementById('db-pending-jobs').innerText = pendingJobsCount;

        if(document.getElementById('db-total-paid')) document.getElementById('db-total-paid').innerText = "?" + totalPaid.toFixed(2);

        if(document.getElementById('db-total-due')) document.getElementById('db-total-due').innerText = "?" + totalDue.toFixed(2);

        if(document.getElementById('db-total-loss')) document.getElementById('db-total-loss').innerText = "?" + totalLoss.toFixed(2);

        if(document.getElementById('db-total-advance')) document.getElementById('db-total-advance').innerText = "?" + totalAdvance.toFixed(2);

    }



    function addCustomer(){

        let nameEl = document.getElementById('name');

        let mobileEl = document.getElementById('mobile');

        if(!nameEl || !mobileEl) return;



        let name = nameEl.value.trim();

        let mobile = mobileEl.value.trim();

        let email = document.getElementById('email')?.value.trim() || '';

        let shop = document.getElementById('shop')?.value.trim() || '';

        let address = document.getElementById('address')?.value.trim() || '';



        if(name === '' || mobile === ''){

            alert("Customer Name and Mobile Number are required.");

            return;

        }



        if(editCustIndex !== null) {

            let oldName = customers[editCustIndex].name;

            if (customers.some((c, i) => i !== editCustIndex && c.name.toLowerCase() === name.toLowerCase())) {

                alert("This Customer Name is already registered!");

                return;

            }

            

            if(oldName !== name) {

                jobs.forEach(job => {

                    if(job.cname === oldName) job.cname = name;

                });

            }



            customers[editCustIndex] = {name, mobile, email, shop, address};

            editCustIndex = null;

            if(document.getElementById('cust-form-title')) document.getElementById('cust-form-title').innerText = "?? New Customer Entry";

            if(document.getElementById('btn-submit-cust')) document.getElementById('btn-submit-cust').innerText = "? Add Customer";

            alert("Customer data updated successfully!");

        } else {

            if (customers.some(c => c.name.toLowerCase() === name.toLowerCase())) {

                alert("This Customer Name is already registered!");

                return;

            }

            customers.push({name, mobile, email, shop, address});

            alert(`Success: ${name} added!`);

        }



        saveState();

        refreshAllViews();

        clearCustomerForm();

    }



    function editCustomer(index) {

        let item = customers[index];

        if(!item) return;

        editCustIndex = index;



        if(document.getElementById("name")) document.getElementById("name").value = item.name;

        if(document.getElementById("mobile")) document.getElementById("mobile").value = item.mobile;

        if(document.getElementById("email")) document.getElementById("email").value = item.email || '';

        if(document.getElementById("shop")) document.getElementById("shop").value = item.shop || '';

        if(document.getElementById("address")) document.getElementById("address").value = item.address || '';

        if(document.getElementById('cust-form-title')) document.getElementById('cust-form-title').innerText = '?? Edit Mode (Updating Customer)';

        if(document.getElementById('btn-submit-cust')) document.getElementById('btn-submit-cust').innerText = '?? Save Changes';

        document.querySelector('#customer .form-card')?.scrollIntoView({ behavior: 'smooth' });
     }

    function deleteCustomer(index){

        let customer = customers[index];

        if(!customer) return;

        if(confirm(`Delete ${customer.name} and all related job entries?`)){

            customers.splice(index, 1);

            jobs = jobs.filter(job => job.cname !== customer.name);

            if(editCustIndex === index) {

                editCustIndex = null;

                clearCustomerForm();

            }

            saveState();

            refreshAllViews();

        }

    }



    function renderCustomers(){

        let table = document.getElementById('table');

        if(!table) return;

        table.innerHTML = "";

        

        if(customers.length === 0) {

            table.innerHTML = `<tr><td colspan="7" style="text-align:center; color:#999;">No active clients.</td></tr>`;

            return;

        }



        let now = new Date();



        customers.forEach((c, i) => {

            let customerJobs = jobs.filter(j => j.cname === c.name && !j.paymentHistory);

            let statusText = "Inactive";

            let statusClass = "status-inactive";

            let lastActivityText = "No Activity";



            if(customerJobs.length > 0) {

                let timestamps = customerJobs.map(j => new Date(j.date).getTime());

                let maxTimestamp = Math.max(...timestamps);

                let lastJobDate = new Date(maxTimestamp);

                

                lastActivityText = lastJobDate.toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });

                

                let diffTime = Math.abs(now - lastJobDate);

                let diffDays = diffTime / (1000 * 60 * 60 * 24);



                if(diffDays <= 1) { 

                    statusText = "Active Today";

                    statusClass = "status-daily"; 

                } else if(diffDays <= 30) {

                    statusText = "Active (<30d)";

                    statusClass = "status-daily"; 

                } else if(diffDays <= 365) {

                    statusText = "Infrequency";

                    statusClass = "status-monthly"; 

                } else {

                    statusText = "Inactive";

                    statusClass = "status-inactive"; 

                }

            }



            table.innerHTML += `<tr>

                <td><strong>${escapeHtml(c.name)}</strong></td>

                <td>${escapeHtml(c.mobile)}</td>

                <td>${escapeHtml(c.email || '—')}</td>

                <td>${escapeHtml(c.shop || '—')}</td>

                <td>

                    <button class="btn-edit" onclick="editCustomer(${i})">?? Edit</button>
                    <button class="btn-danger" onclick="deleteCustomer(${i})">?? Delete</button>

                </td>

                <td><span class="badge ${statusClass}">${statusText}</span></td>

                <td><code>${lastActivityText}</code></td>

            </tr>`;

        });

    }



    function clearCustomerForm() {

        if(document.getElementById('name')) document.getElementById('name').value = '';

        if(document.getElementById('mobile')) document.getElementById('mobile').value = '';

        if(document.getElementById('email')) document.getElementById('email').value = '';

        if(document.getElementById('shop')) document.getElementById('shop').value = '';

        if(document.getElementById('address')) document.getElementById('address').value = '';

    }



    function toggleCustomerDropdown() {

        let box = document.getElementById("customerList");

        if(!box) return;

        if(customers.length === 0) {

            alert("Please register a customer first!");

            return;

        }

        box.innerHTML = "";

        customers.forEach(c => {

            let div = document.createElement("div");

            div.className = "dropdown-item";

            div.innerText = c.name + (c.shop ? ` (${c.shop})` : "");

            div.onclick = function(e){

                e.stopPropagation();

                if(document.getElementById("cname")) document.getElementById("cname").value = c.name;

                box.style.display = "none";

            };

            box.appendChild(div);

        });

        box.style.display = box.style.display === "none" ? "block" : "none";

    }



    function setDate(){

        let today = new Date();

        let yyyy = today.getFullYear();

        let mm = String(today.getMonth() + 1).padStart(2, '0');

        let dd = String(today.getDate()).padStart(2, '0');

        

        let dateField = document.getElementById('jobDate');

        if(dateField) {

            dateField.value = `${yyyy}-${mm}-${dd}`;

            dateField.readOnly = true;

            dateField.style.background = "#eef2f5";

        }

    }



    function addJob(){

        let cname = document.getElementById('cname')?.value;

        let description = document.getElementById('description')?.value.trim() || '';

        let size = document.getElementById('size')?.value.trim() || '';

        let price = parseFloat(document.getElementById('price')?.value);

        let status = document.getElementById('status')?.value || 'Paid';

        let selectedDateVal = document.getElementById('jobDate')?.value;



        if(!cname || cname === ''){

            alert("Please assign a valid Customer.");

            return;

        }

        if(isNaN(price) || price < 0) {

            alert("Please assign a valid numeric price.");

            return;

        }



        let targetDate = new Date(selectedDateVal);

        if(isNaN(targetDate.getTime())) {

            alert("Please select a valid entry date.");

            return;

        }

        let curTime = new Date();

        targetDate.setHours(curTime.getHours(), curTime.getMinutes(), curTime.getSeconds());

        let finalIsoStr = targetDate.toISOString();



        if(editIndex !== null) {

            jobs[editIndex] = {

                ...jobs[editIndex], 

                cname,

                description: description || '—',

                size: size || '—',

                price: price,

                status: status,

                date: finalIsoStr, 

                payDate: status === 'Paid' ? (jobs[editIndex].payDate || finalIsoStr) : null

            };

            if(!jobs[editIndex].id) jobs[editIndex].id = createId('job');

            editIndex = null;

            if(document.getElementById('form-title')) document.getElementById('form-title').innerText = "?? Order & Data Entry";

            if(document.getElementById('btn-submit-job')) document.getElementById('btn-submit-job').innerText = "? Add Entry";

            alert("Entry updated successfully!");

        } else {

            jobs.push({

                id: createId('job'),

                date: finalIsoStr,

                cname,

                description: description || '—',

                size: size || '—',

                price: price,

                status: status,

                payDate: status === 'Paid' ? finalIsoStr : null

            });

            alert("New entry added successfully!");

        }



        saveState();

        updateDashboardMetrics();

        filterJobsTable();

        generateFullReport();

        generateMonthlyMatrixReport();

        clearJobForm();

        setDate();

    }



    function editJob(index) {

        let item = jobs[index];

        editIndex = index;



        if(document.getElementById("cname")) document.getElementById("cname").value = item.cname;

        if(document.getElementById("description")) document.getElementById("description").value = item.description === '—' ? '' : item.description;

        if(document.getElementById("size")) document.getElementById("size").value = item.size === '—' ? '' : item.size;

        if(document.getElementById("price")) document.getElementById("price").value = item.price;

        if(document.getElementById("status")) document.getElementById("status").value = item.status;

        

        let itemDate = new Date(item.date);

        let yyyy = itemDate.getFullYear();

        let mm = String(itemDate.getMonth() + 1).padStart(2, '0');

        let dd = String(itemDate.getDate()).padStart(2, '0');

        

        let dateField = document.getElementById("jobDate");

        if(dateField) {

            dateField.value = `${yyyy}-${mm}-${dd}`;

            dateField.readOnly = false;

            dateField.style.background = "#fff";

        }



        if(document.getElementById('form-title')) document.getElementById('form-title').innerText = "?? Edit Mode (Updating Entry)";

        if(document.getElementById('btn-submit-job')) document.getElementById('btn-submit-job').innerText = "?? Save Changes";

        

        document.querySelector('#dataentry .form-card')?.scrollIntoView({ behavior: 'smooth' });

    }



    function clearJobForm() {

        if(document.getElementById('cname')) document.getElementById('cname').value = '';

        if(document.getElementById('description')) document.getElementById('description').value = '';

        if(document.getElementById('size')) document.getElementById('size').value = '';

        if(document.getElementById('price')) document.getElementById('price').value = '';

        if(document.getElementById('status')) document.getElementById('status').value = 'Paid';

    }



    function populatePivotDropdown() {

        let select = document.getElementById("filterCustSelect");

        if (!select) return;

        

        let currentSelection = select.value || "All";

        select.innerHTML = "";

        let allOption = document.createElement("option");

        allOption.value = "All";

        allOption.textContent = "(All Registered Clients)";

        select.appendChild(allOption);

        

        customers.forEach(c => {

            let option = document.createElement("option");

            option.value = c.name;

            option.textContent = c.name;

            select.appendChild(option);

        });

        

        select.value = customers.some(c => c.name === currentSelection) ? currentSelection : "All";

    }



    function filterJobsTable() {

        let clientFilter = document.getElementById("filterCustSelect")?.value || "All";

        if(clientFilter === "") clientFilter = "All";

        let fromDateVal = document.getElementById("filterFromDate")?.value || "";

        let toDateVal = document.getElementById("filterToDate")?.value || "";

        let statusFilter = document.getElementById("filterStatus")?.value || "All";

        

        let table = document.getElementById('jobTable');

        if (!table) return;

        table.innerHTML = "";

        

        let startTimestamp = fromDateVal ? new Date(fromDateVal).setHours(0,0,0,0) : null;

        let endTimestamp = toDateVal ? new Date(toDateVal).setHours(23,59,59,999) : null;

        

        let filteredCount = 0;

        

        for (let i = jobs.length - 1; i >= 0; i--) {

           if (jobs[i].paymentHistory) continue;

            let j = jobs[i];

            let jDate = new Date(j.date);

            let jTime = jDate.getTime();

            

            if (clientFilter !== "All" && j.cname !== clientFilter) continue;

            if (statusFilter !== "All" && j.status !== statusFilter) continue;

            if (startTimestamp && jTime < startTimestamp) continue;

            if (endTimestamp && jTime > endTimestamp) continue;

            

            filteredCount++;

            let jobDateFormatted = jDate.toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });

            let statusStyleClass = j.status === "Paid" ? "paid" : (j.status === "Cancelled" ? "loss-text" : "unpaid");

            

            table.innerHTML += `

            <tr>

                <td>${jobDateFormatted}</td>

                <td><strong>${escapeHtml(j.cname)}</strong></td>

                <td>${escapeHtml(j.description)}</td>

                <td>${escapeHtml(j.size)}</td>

                <td>?${parseFloat(j.price).toFixed(2)}</td>

                <td class="${statusStyleClass}">${escapeHtml(j.status)}${j.status === "Unpaid" && getPaidAmountForJob(j) > 0 ? ` (Due ?${getOutstandingAmount(j).toFixed(2)})` : ""}</td>

                <td>

                    <button class="btn-edit" onclick="editJob(${i})">?? Edit</button>
                    <button class="btn-danger" onclick="deleteJob(${i})">?? Delete</button>

                </td>

            </tr>`;

        }

        

        if (filteredCount === 0) {

            table.innerHTML = `<tr><td colspan="7" style="text-align:center; color:#999;">No entries match filters.</td></tr>`;

        }

    }



    function resetJobFilters() {

        if(document.getElementById("filterCustSelect")) document.getElementById("filterCustSelect").value = "All";

        if(document.getElementById("filterFromDate")) document.getElementById("filterFromDate").value = "";

        if(document.getElementById("filterToDate")) document.getElementById("filterToDate").value = "";

        if(document.getElementById("filterStatus")) document.getElementById("filterStatus").value = "All";

        filterJobsTable();

    }



    function generateFullReport() {

        let tableBody = document.getElementById("reportTable");

        if (!tableBody) return;

        tableBody.innerHTML = "";

        

        if (customers.length === 0) {

            tableBody.innerHTML = `<tr><td colspan="10" style="text-align:center; color:#999;">Register clients to see revenue data.</td></tr>`;

            return;

        }

        

        let now = new Date();

        let dynamicMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();

        

        let standardDayMs = 24 * 60 * 60 * 1000;

        let limit7d = dynamicMidnight - (7 * standardDayMs);

        let limit30d = dynamicMidnight - (30 * standardDayMs);

        let limit365d = dynamicMidnight - (365 * standardDayMs);

        

        customers.forEach(c => {

            let clientJobs = jobs.filter(j => j.cname === c.name);

            let revToday = 0, rev7d = 0, rev30d = 0, rev365d = 0;

            let totalValueDone = 0, totalPaid = 0, totalUnpaid = 0, totalCancelled = 0;

            

            clientJobs.forEach(j => {

                let val = parseFloat(j.price) || 0;

                let incomeTime = new Date(j.payDate || j.date).getTime();

                

                if (isIncomeEntry(j)) totalPaid += getIncomeAmount(j);

                else if (j.status === "Unpaid") totalUnpaid += getOutstandingAmount(j);

                else if (j.status === "Cancelled") totalCancelled += val;

                

                if (isIncomeEntry(j)) {

    val = getIncomeAmount(j);

    totalValueDone += val;

    if (incomeTime >= dynamicMidnight) revToday += val;
    if (incomeTime >= limit7d) rev7d += val;
    if (incomeTime >= limit30d) rev30d += val;
    if (incomeTime >= limit365d) rev365d += val;
}
            });

            

            // Generate WhatsApp button conditional code snippet

            let whatsappBtn = "";

            if (totalUnpaid > 0) {

                whatsappBtn = `<button class="btn-whatsapp" onclick="sendWhatsAppReminder(decodeFromClick('${encodeForClick(c.name)}'), decodeFromClick('${encodeForClick(c.mobile)}'), ${totalUnpaid})">?? WhatsApp</button>`;

            }

            

            tableBody.innerHTML += `

            <tr>

                <td><strong>${escapeHtml(c.name)}</strong></td>

                <td>?${revToday.toFixed(2)}</td>

                <td>?${rev7d.toFixed(2)}</td>

                <td>?${rev30d.toFixed(2)}</td>

                <td>?${rev365d.toFixed(2)}</td>

                <td><strong>?${totalValueDone.toFixed(2)}</strong></td>

                <td class="paid">?${totalPaid.toFixed(2)}</td>

                <td class="unpaid">?${totalUnpaid.toFixed(2)}</td>

                <td style="color:#e74c3c; text-decoration:line-through;">?${totalCancelled.toFixed(2)}</td>

                <td style="text-align: center;">

                    <button class="btn-invoice" onclick="openInvoiceModal(decodeFromClick('${encodeForClick(c.name)}'))">?? Settle</button>

                    <button class="btn-history" onclick="openHistoryModal(decodeFromClick('${encodeForClick(c.name)}'))">?? History</button>

                    ${whatsappBtn}

                </td>

            </tr>`;

        });

    }



    // New WhatsApp Action logic block

    function sendWhatsAppReminder(customerName, mobileNumber, dueAmount) {

        let cleanNumber = mobileNumber.replace(/\D/g, '');

        // If 10 digits (Standard Indian Mobile format without country code prefix), auto-prepend 91

        if (cleanNumber.length === 10) {

            cleanNumber = '91' + cleanNumber;

        }

        if (cleanNumber.length < 11) {

            alert("Please add a valid mobile number before sending WhatsApp reminder.");

            return;

        }

        

        let messageText = `Dear *${customerName}*,\n\nThis is a friendly reminder from *Design Studio*. Your outstanding ledger balance statement shows a pending due amount of *?${dueAmount.toFixed(2)}*.\n\nKindly process the payment clear balance at your earliest convenience.\n\nThank you for choosing us! ?`;

        let encodedMessage = encodeURIComponent(messageText);

        let waApiUrl = `https://wa.me/${cleanNumber}?text=${encodedMessage}`;

        

        window.open(waApiUrl, '_blank');

    }



    function openInvoiceModal(customerName) {

        currentActiveModalCustomer = customerName;

        if(document.getElementById("modalCustomerName")) document.getElementById("modalCustomerName").innerText = customerName;

        

        let tbody = document.getElementById("modalInvoiceTableBody");

        if(!tbody) return;

        tbody.innerHTML = "";

        

        let dueJobs = jobs.filter(j => j.cname === customerName && j.status === "Unpaid" && getOutstandingAmount(j) > 0);

        currentCalculatedTotalDue = 0;

        

        dueJobs.forEach(j => {

            let originalVal = parseFloat(j.price) || 0;

            let paidVal = getPaidAmountForJob(j);

            let val = getOutstandingAmount(j);

            currentCalculatedTotalDue += val;

            let jobDateStr = new Date(j.date).toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });

            

            tbody.innerHTML += `

            <tr>

                <td style="padding:10px;">${jobDateStr}</td>

                <td style="padding:10px;">${escapeHtml(j.description)}</td>

                <td style="padding:10px;">${escapeHtml(j.size)}</td>

                <td style="padding:10px; text-align:right;" class="unpaid">?${val.toFixed(2)}${paidVal > 0 ? `<br><small style="color:#64748b;">Original ?${originalVal.toFixed(2)} | Paid ?${paidVal.toFixed(2)}</small>` : ""}</td>

            </tr>`;

        });

        

        if (dueJobs.length === 0) {

            tbody.innerHTML = `<tr><td colspan="4" style="text-align:center; padding:20px; color:#2e7d32; font-weight:600;">?? No pending invoices. You can still record an advance payment below.</td></tr>`;

        }

        if(document.getElementById("cashReceivedInput")) document.getElementById("cashReceivedInput").disabled = false;

        

        if(document.getElementById("modalTotalDueAmt")) document.getElementById("modalTotalDueAmt").innerText = "?" + currentCalculatedTotalDue.toFixed(2);

        if(document.getElementById("cashReceivedInput")) {

            document.getElementById("cashReceivedInput").value = "";

            document.getElementById("cashReceivedInput").placeholder = currentCalculatedTotalDue > 0 ? "Enter Cash Received Amount" : "Enter Advance Payment Amount";

        }

        

        let printBtn = document.getElementById("modalPrintBtn");

        if(printBtn) {

            printBtn.onclick = function() { window.print(); };

        }

        

        if(document.getElementById("invoiceModal")) document.getElementById("invoiceModal").style.display = "flex";

    }



    function closeInvoiceModal() {

        if(document.getElementById("invoiceModal")) document.getElementById("invoiceModal").style.display = "none";

        generateFullReport();

    }



    function processSmartSettlement() {

    let cashPaid = parseFloat(document.getElementById("cashReceivedInput")?.value);

    if (isNaN(cashPaid) || cashPaid <= 0) {
        alert("Please enter valid amount.");
        return;
    }

    let remainingCash = cashPaid;
    let nowIsoString = new Date().toISOString();

    // FIFO Order
    let customerJobs = jobs
        .filter(j => j.cname === currentActiveModalCustomer && j.status === "Unpaid" && getOutstandingAmount(j) > 0)
        .sort((a, b) => new Date(a.date) - new Date(b.date));

    customerJobs.forEach(j => {

        if (remainingCash <= 0) return;

        let dueAmount = getOutstandingAmount(j);

        // FULL PAYMENT
        if (remainingCash >= dueAmount) {

            jobs.push({
                id: createId('pay'),
                relatedJobId: j.id,
                date: nowIsoString,
                cname: currentActiveModalCustomer,
                description: j.description,
                size: j.size,
                price: dueAmount,
                status: "Paid",
                payDate: nowIsoString,
                paymentHistory: true
            });

            j.status = "Paid";
            j.payDate = nowIsoString;

            remainingCash -= dueAmount;
        }

        // PARTIAL PAYMENT
        else {

            let paidPart = remainingCash;

            // CREATE ONLY PAYMENT HISTORY ENTRY
            jobs.push({
                id: createId('pay'),
                relatedJobId: j.id,
                date: nowIsoString,
                cname: currentActiveModalCustomer,
                description: j.description + " (Partial Paid)",
                size: j.size,
                price: paidPart,
                status: "Paid",
                payDate: nowIsoString,
                paymentHistory: true
            });

            remainingCash = 0;
        }
    });

    if (remainingCash > 0) {

        jobs.push({
            id: createId('pay'),
            relatedJobId: null,
            date: nowIsoString,
            cname: currentActiveModalCustomer,
            description: "Advance / Extra Payment",
            size: "—",
            price: remainingCash,
            status: "Paid",
            payDate: nowIsoString,
            paymentHistory: true
        });

        remainingCash = 0;

    }

    saveState();

    refreshAllViews();

    alert(currentCalculatedTotalDue > 0 ? "Settlement completed successfully!" : "Advance payment recorded successfully!");

    openInvoiceModal(currentActiveModalCustomer);
    syncToGoogleSheets();
}


    function openHistoryModal(customerName) {

        if(document.getElementById("historyCustomerName")) document.getElementById("historyCustomerName").innerText = `${customerName} - History`;

        let tbody = document.getElementById("historyTableBody");

        if(!tbody) return;

        tbody.innerHTML = "";

        

        let paidJobs = jobs.filter(j => j.cname === customerName && isIncomeEntry(j));

        paidJobs.sort((a, b) => new Date(b.payDate || b.date) - new Date(a.payDate || a.date));

        

        paidJobs.forEach(j => {

            let clearanceDate = new Date(j.payDate || j.date).toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });

            tbody.innerHTML += `

            <tr>

                <td style="padding:10px;"><code>${clearanceDate}</code></td>

                <td style="padding:10px;">${escapeHtml(j.description)}</td>

                <td style="padding:10px;">${escapeHtml(j.size)}</td>

                <td style="padding:10px; text-align:right;" class="paid">?${getIncomeAmount(j).toFixed(2)}</td>

            </tr>`;

        });

        

        if (paidJobs.length === 0) {

            tbody.innerHTML = `<tr><td colspan="4" style="text-align:center; padding:20px; color:#999;">No previous payment records found.</td></tr>`;

        }

        

        if(document.getElementById("historyModal")) document.getElementById("historyModal").style.display = "flex";

    }



    function closeHistoryModal() {

        if(document.getElementById("historyModal")) document.getElementById("historyModal").style.display = "none";

    }


    function deleteJob(index){

        if(!jobs[index]) return;

        if(confirm("Delete this entry?")){

            let deletedJob = jobs[index];

            jobs = jobs.filter((job, i) => i !== index && job.relatedJobId !== deletedJob.id);

            saveState();

            refreshAllViews();

        }

    }
    


    function clearSystemData(){
        if(confirm('Clear all saved customers and jobs? This cannot be undone.')){
            customers = [];
            jobs = [];
            editIndex = null;
            editCustIndex = null;
            localStorage.removeItem('crm_customers');
            localStorage.removeItem('crm_jobs');
            refreshAllViews();
            clearCustomerForm();
            clearJobForm();
            setDate();
            alert('All saved data cleared.');
        }
    }

    function exportJobsToCSV() {
        if (jobs.length === 0) {
            alert("No data to export.");
            return;
        }
        let csvRows = [["Date", "Customer", "Description", "Size", "Price", "Status"]];
        jobs.filter(j => !j.paymentHistory).forEach(j => {
            csvRows.push([
                new Date(j.date).toLocaleDateString(),
                `"${j.cname}"`,
                `"${j.description}"`,
                `"${j.size}"`,
                j.price,
                j.status
            ]);
        });
        let csvContent = "data:text/csv;charset=utf-8," + csvRows.map(e => e.join(",")).join("\n");
        let encodedUri = encodeURI(csvContent);
        let link = document.createElement("a");
        link.setAttribute("href", encodedUri);
        link.setAttribute("download", `DesignStudio_Ledger_${new Date().toISOString().split('T')[0]}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    function backupData() {
        const data = { customers, jobs };
        const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `crm_backup_${new Date().toISOString().slice(0,10)}.json`;
        a.click();
        URL.revokeObjectURL(url);
    }

    function restoreData(input) {
        const file = input.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = JSON.parse(e.target.result);
                if (data.customers && data.jobs) {
                    if (confirm("Warning: Restoring data will overwrite your current database. Do you want to proceed?")) {
                        customers = data.customers;
                        jobs = data.jobs;
                        if(ensureJobIds(jobs)) {
                            localStorage.setItem('crm_jobs', JSON.stringify(jobs));
                        }
                        saveState();
                        refreshAllViews();
                        alert("Database restored successfully!");
                    }
                } else {
                    alert("Error: The selected file is not a valid Design Studio backup.");
                }
            } catch (err) {
                alert("Error: Failed to read the backup file.");
            }
            input.value = '';
        };
        reader.readAsText(file);
    }

    async function syncToGoogleSheets(event = null) {
        let btn = (event && event.currentTarget) ? event.currentTarget : null;
        let originalText = btn ? btn.innerText : "";

        if (btn) {
            btn.innerText = "? Syncing...";
            btn.disabled = true;
        }

        try {
            if (isAppsScriptContext()) {
                saveStateToGoogleSheets();
                if (event) alert("Google Sheet-? ???? ??????? ????? ?????!");
            } else {
                const scriptURL = 'https://script.google.com/macros/s/AKfycbwLl5nVHJp_-g3nzzBWTjwg_bMdV1c-KSaIAyFGaypyB5lhx_0aYYb-KHq9iYqYL0GHgg/exec';
                await fetch(scriptURL, {
                    method: 'POST',
                    mode: 'no-cors',
                    body: JSON.stringify({ customers, jobs })
                });
                if (event) alert("Google Sheet-? ???? ??????? ????? ?????!");
            }
        } catch (error) {
            console.error('Sync error:', error);
            if (event) alert("????? ???? ?????? ?????? ????? ??? ?????");
        } finally {
            if (btn) {
                btn.innerText = originalText;
                btn.disabled = false;
            }
        }
    }

    function toggleDarkMode() {
        document.body.classList.toggle('dark-mode');
        const isDark = document.body.classList.contains('dark-mode');
        localStorage.setItem('theme', isDark ? 'dark' : 'light');
        updateThemeButton();
    }

    function updateThemeButton() {
        const btn = document.getElementById('theme-toggle');
        if (btn) btn.innerText = document.body.classList.contains('dark-mode') ? "?? Light Mode" : "?? Dark Mode";
    }

    function initTheme() {
        const savedTheme = localStorage.getItem('theme');
        if (savedTheme === 'dark') {
            document.body.classList.add('dark-mode');
        }
        updateThemeButton();
    }

