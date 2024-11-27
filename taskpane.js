Office.onReady(async (info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById('create-task').onclick = createPlannerTask;
        await loadPlannerLists();
        await loadRelatedTasks();
    }
});

async function loadPlannerLists() {
    const plannerLists = await graphAPI('/me/planner/plans');
    const dropdown = document.getElementById('planner-list');
    plannerLists.value.forEach(plan => {
        const option = document.createElement('option');
        option.value = plan.id;
        option.textContent = plan.title;
        dropdown.appendChild(option);
    });
}

async function createPlannerTask() {
    const planId = document.getElementById('planner-list').value;
    const email = Office.context.mailbox.item;

    const task = {
        planId: planId,
        title: email.subject,
        assignments: {},
        details: {
            description: `Von E-Mail: ${email.internetMessageId}`
        }
    };

    const createdTask = await graphAPI(`/planner/tasks`, 'POST', task);
    email.body.setAsync(`<p>Planner Task ID: ${createdTask.id}</p>`, { coercionType: Office.CoercionType.Html });
    alert(`Aufgabe erstellt: ${createdTask.title}`);
}

async function loadRelatedTasks() {
    const email = Office.context.mailbox.item;
    const tasks = await graphAPI('/me/planner/tasks');
    const relatedTasks = tasks.value.filter(task => task.details.description.includes(email.internetMessageId));

    const list = document.getElementById('related-tasks');
    list.innerHTML = '';
    relatedTasks.forEach(task => {
        const li = document.createElement('li');
        li.textContent = task.title;
        list.appendChild(li);
    });
}

async function graphAPI(endpoint, method = 'GET', body = null) {
    const token = await Office.context.auth.getAccessTokenAsync();
    const headers = {
        'Authorization': `Bearer ${token.value}`,
        'Content-Type': 'application/json'
    };

    const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
        method,
        headers,
        body: body ? JSON.stringify(body) : null
    });

    if (!response.ok) {
        throw new Error(`Graph API call failed: ${response.statusText}`);
    }

    return response.json();
}
