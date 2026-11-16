/**
 * TaskGravity - Practical 02: Interactive Client-Side Application
 * JSON and Arrays Data Storage, Comprehensive Client-Side Validation, and Dynamic DOM Manipulation
 */

// Initial Seed Tasks (JSON Array)
const DEFAULT_TASKS = [
  {
    id: "tg-101",
    title: "Implement Responsive Flexbox Grid",
    category: "Frontend",
    priority: "High",
    dueDate: "2026-11-20",
    estimatedHours: 4,
    description: "Audit breakpoints across mobile and desktop to eliminate horizontal overflow.",
    status: "in-progress",
    createdAt: "2026-11-15T09:00:00Z"
  },
  {
    id: "tg-102",
    title: "Configure Express Validator Middleware",
    category: "Backend",
    priority: "Urgent",
    dueDate: "2026-11-22",
    estimatedHours: 6,
    description: "Write schema sanitation and validation rules for task POST endpoints.",
    status: "pending",
    createdAt: "2026-11-15T11:30:00Z"
  },
  {
    id: "tg-103",
    title: "Draft TaskGravity Mongoose Schema",
    category: "Database",
    priority: "Medium",
    dueDate: "2026-11-24",
    estimatedHours: 3,
    description: "Define indices, enum constraints, and timestamps for MongoDB Atlas collection.",
    status: "completed",
    createdAt: "2026-11-14T14:15:00Z"
  },
  {
    id: "tg-104",
    title: "Postman Integration Test Suite",
    category: "Testing",
    priority: "Low",
    dueDate: "2026-11-26",
    estimatedHours: 5,
    description: "Create collection runners verifying 200, 201, 400, and 404 response codes.",
    status: "pending",
    createdAt: "2026-11-16T08:00:00Z"
  }
];

class TaskGravityApp {
  constructor() {
    this.storageKey = "TaskGravity_Tasks_v2";
    this.tasks = this.loadTasks();
    this.activeFilterStatus = "all";
    this.activeFilterCategory = "all";
    this.searchQuery = "";
    this.sortBy = "dueDate";

    this.initDOMElements();
    this.bindEvents();
    this.render();
  }

  // Load JSON data from localStorage or fallback to default array
  loadTasks() {
    try {
      const stored = localStorage.getItem(this.storageKey);
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed) && parsed.length > 0) {
          return parsed;
        }
      }
    } catch (e) {
      console.warn("Could not read from localStorage, using default seed:", e);
    }
    return [...DEFAULT_TASKS];
  }

  // Persist JSON array to localStorage
  saveTasks() {
    try {
      localStorage.setItem(this.storageKey, JSON.stringify(this.tasks));
    } catch (e) {
      console.error("Failed to persist tasks to storage:", e);
    }
  }

  initDOMElements() {
    // Form & Inputs
    this.form = document.getElementById("taskForm");
    this.titleInput = document.getElementById("taskTitle");
    this.categorySelect = document.getElementById("taskCategory");
    this.prioritySelect = document.getElementById("taskPriority");
    this.dueDateInput = document.getElementById("taskDueDate");
    this.hoursInput = document.getElementById("taskHours");
    this.descInput = document.getElementById("taskDesc");

    // Feedback elements
    this.errorTitle = document.getElementById("errorTitle");
    this.errorCategory = document.getElementById("errorCategory");
    this.errorPriority = document.getElementById("errorPriority");
    this.errorDueDate = document.getElementById("errorDueDate");
    this.errorHours = document.getElementById("errorHours");
    this.errorDesc = document.getElementById("errorDesc");

    // Task List & Controls
    this.taskListContainer = document.getElementById("taskList");
    this.searchInput = document.getElementById("searchInput");
    this.statusFilterSelect = document.getElementById("statusFilter");
    this.categoryFilterSelect = document.getElementById("categoryFilter");
    this.sortBySelect = document.getElementById("sortBy");
    this.resetDataBtn = document.getElementById("resetDataBtn");

    // Stats
    this.statTotal = document.getElementById("statTotal");
    this.statCompleted = document.getElementById("statCompleted");
    this.statInProgress = document.getElementById("statInProgress");
    this.statPending = document.getElementById("statPending");
    this.statVelocity = document.getElementById("statVelocity");

    // Toast
    this.toastBanner = document.getElementById("toastBanner");

    // Set min date for date picker to today
    const todayStr = new Date().toISOString().split("T")[0];
    this.dueDateInput.min = todayStr;
  }

  bindEvents() {
    // Form Submission
    this.form.addEventListener("submit", (e) => this.handleFormSubmit(e));

    // Real-time input validation on blur / input
    this.titleInput.addEventListener("input", () => this.validateField("title"));
    this.categorySelect.addEventListener("change", () => this.validateField("category"));
    this.prioritySelect.addEventListener("change", () => this.validateField("priority"));
    this.dueDateInput.addEventListener("change", () => this.validateField("dueDate"));
    this.hoursInput.addEventListener("input", () => this.validateField("hours"));
    this.descInput.addEventListener("input", () => this.validateField("description"));

    // Filters and Search
    this.searchInput.addEventListener("input", (e) => {
      this.searchQuery = e.target.value.toLowerCase().trim();
      this.renderTaskList();
    });

    this.statusFilterSelect.addEventListener("change", (e) => {
      this.activeFilterStatus = e.target.value;
      this.renderTaskList();
    });

    this.categoryFilterSelect.addEventListener("change", (e) => {
      this.activeFilterCategory = e.target.value;
      this.renderTaskList();
    });

    this.sortBySelect.addEventListener("change", (e) => {
      this.sortBy = e.target.value;
      this.renderTaskList();
    });

    if (this.resetDataBtn) {
      this.resetDataBtn.addEventListener("click", () => {
        if (confirm("Reset task board to initial demonstration dataset?")) {
          this.tasks = [...DEFAULT_TASKS];
          this.saveTasks();
          this.render();
          this.showToast("Dataset reset to factory defaults.");
        }
      });
    }
  }

  // Comprehensive Input Validation
  validateField(field) {
    let isValid = true;

    if (field === "title" || field === "all") {
      const val = this.titleInput.value.trim();
      if (!val) {
        this.setError(this.titleInput, this.errorTitle, "Task Title is required.");
        isValid = false;
      } else if (val.length < 3) {
        this.setError(this.titleInput, this.errorTitle, "Task Title must be at least 3 characters.");
        isValid = false;
      } else if (val.length > 80) {
        this.setError(this.titleInput, this.errorTitle, "Task Title cannot exceed 80 characters.");
        isValid = false;
      } else {
        this.clearError(this.titleInput, this.errorTitle);
      }
    }

    if (field === "category" || field === "all") {
      const val = this.categorySelect.value;
      const validCategories = ["Frontend", "Backend", "Database", "DevOps", "Testing", "UI/UX"];
      if (!val || !validCategories.includes(val)) {
        this.setError(this.categorySelect, this.errorCategory, "Please select a valid architectural category.");
        isValid = false;
      } else {
        this.clearError(this.categorySelect, this.errorCategory);
      }
    }

    if (field === "priority" || field === "all") {
      const val = this.prioritySelect.value;
      const validPriorities = ["Low", "Medium", "High", "Urgent"];
      if (!val || !validPriorities.includes(val)) {
        this.setError(this.prioritySelect, this.errorPriority, "Please select an operational priority level.");
        isValid = false;
      } else {
        this.clearError(this.prioritySelect, this.errorPriority);
      }
    }

    if (field === "dueDate" || field === "all") {
      const val = this.dueDateInput.value;
      if (!val) {
        this.setError(this.dueDateInput, this.errorDueDate, "Target deadline is required.");
        isValid = false;
      } else {
        const selected = new Date(val);
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        if (isNaN(selected.getTime())) {
          this.setError(this.dueDateInput, this.errorDueDate, "Invalid deadline date format.");
          isValid = false;
        } else if (selected < today) {
          this.setError(this.dueDateInput, this.errorDueDate, "Deadline cannot be scheduled in the past.");
          isValid = false;
        } else {
          this.clearError(this.dueDateInput, this.errorDueDate);
        }
      }
    }

    if (field === "hours" || field === "all") {
      const val = parseFloat(this.hoursInput.value);
      if (isNaN(val) || val <= 0) {
        this.setError(this.hoursInput, this.errorHours, "Estimated effort must be a positive number.");
        isValid = false;
      } else if (val > 100) {
        this.setError(this.hoursInput, this.errorHours, "Estimated effort cannot exceed 100 hours per unit.");
        isValid = false;
      } else {
        this.clearError(this.hoursInput, this.errorHours);
      }
    }

    if (field === "description" || field === "all") {
      const val = this.descInput.value.trim();
      if (val.length > 300) {
        this.setError(this.descInput, this.errorDesc, "Description cannot exceed 300 characters.");
        isValid = false;
      } else {
        this.clearError(this.descInput, this.errorDesc);
      }
    }

    return isValid;
  }

  setError(input, errorEl, message) {
    input.classList.add("is-invalid");
    errorEl.textContent = message;
    errorEl.classList.add("visible");
  }

  clearError(input, errorEl) {
    input.classList.remove("is-invalid");
    errorEl.textContent = "";
    errorEl.classList.remove("visible");
  }

  handleFormSubmit(e) {
    e.preventDefault();

    const isFormValid = this.validateField("all");
    if (!isFormValid) {
      this.showToast("Please correct validation errors before saving.");
      return;
    }

    // Construct new JSON task object
    const newTask = {
      id: "tg-" + Date.now().toString().slice(-4),
      title: this.titleInput.value.trim(),
      category: this.categorySelect.value,
      priority: this.prioritySelect.value,
      dueDate: this.dueDateInput.value,
      estimatedHours: parseFloat(this.hoursInput.value),
      description: this.descInput.value.trim(),
      status: "pending",
      createdAt: new Date().toISOString()
    };

    // Prepend to array
    this.tasks.unshift(newTask);
    this.saveTasks();

    // Reset Form
    this.form.reset();
    this.render();
    this.showToast(`Task "${newTask.title}" successfully created!`);
  }

  toggleTaskStatus(id) {
    const task = this.tasks.find(t => t.id === id);
    if (!task) return;

    // Cycle: pending -> in-progress -> completed -> pending
    if (task.status === "pending") {
      task.status = "in-progress";
      this.showToast(`"${task.title}" marked In Progress.`);
    } else if (task.status === "in-progress") {
      task.status = "completed";
      this.showToast(`"${task.title}" marked Completed.`);
    } else {
      task.status = "pending";
      this.showToast(`"${task.title}" moved to Pending.`);
    }

    this.saveTasks();
    this.render();
  }

  deleteTask(id) {
    const taskIndex = this.tasks.findIndex(t => t.id === id);
    if (taskIndex === -1) return;

    const taskTitle = this.tasks[taskIndex].title;
    if (confirm(`Are you sure you want to delete task: "${taskTitle}"?`)) {
      this.tasks.splice(taskIndex, 1);
      this.saveTasks();
      this.render();
      this.showToast(`Task "${taskTitle}" removed.`);
    }
  }

  getFilteredTasks() {
    return this.tasks.filter(task => {
      // Status filter
      if (this.activeFilterStatus !== "all" && task.status !== this.activeFilterStatus) {
        return false;
      }
      // Category filter
      if (this.activeFilterCategory !== "all" && task.category !== this.activeFilterCategory) {
        return false;
      }
      // Search query
      if (this.searchQuery) {
        const matchesTitle = task.title.toLowerCase().includes(this.searchQuery);
        const matchesDesc = task.description.toLowerCase().includes(this.searchQuery);
        if (!matchesTitle && !matchesDesc) {
          return false;
        }
      }
      return true;
    }).sort((a, b) => {
      if (this.sortBy === "dueDate") {
        return new Date(a.dueDate) - new Date(b.dueDate);
      }
      if (this.sortBy === "title") {
        return a.title.localeCompare(b.title);
      }
      if (this.sortBy === "priority") {
        const weight = { "Urgent": 4, "High": 3, "Medium": 2, "Low": 1 };
        return (weight[b.priority] || 0) - (weight[a.priority] || 0);
      }
      if (this.sortBy === "hours") {
        return b.estimatedHours - a.estimatedHours;
      }
      return new Date(b.createdAt) - new Date(a.createdAt);
    });
  }

  renderStats() {
    const total = this.tasks.length;
    const completed = this.tasks.filter(t => t.status === "completed").length;
    const inProgress = this.tasks.filter(t => t.status === "in-progress").length;
    const pending = this.tasks.filter(t => t.status === "pending").length;
    const velocity = total > 0 ? Math.round((completed / total) * 100) : 0;

    this.statTotal.textContent = total;
    this.statCompleted.textContent = completed;
    this.statInProgress.textContent = inProgress;
    this.statPending.textContent = pending;
    this.statVelocity.textContent = `${velocity}%`;
  }

  renderTaskList() {
    const filtered = this.getFilteredTasks();

    if (filtered.length === 0) {
      this.taskListContainer.innerHTML = `
        <div class="empty-state">
          <div class="empty-state-title">No tasks match current filter criteria</div>
          <p>Adjust your search filters or create a new task using the submission form.</p>
        </div>
      `;
      return;
    }

    this.taskListContainer.innerHTML = filtered.map(task => {
      const isCompleted = task.status === "completed";
      const priorityClass = `badge-${task.priority.toLowerCase()}`;
      const statusClass = `badge-status-${task.status}`;
      const statusLabel = task.status === "in-progress" ? "In Progress" : (task.status === "completed" ? "Completed" : "Pending");

      return `
        <article class="task-item ${isCompleted ? 'completed' : ''}" data-id="${task.id}">
          <div class="task-item-top">
            <label class="task-checkbox-wrap" title="Click to cycle status: Pending -> In Progress -> Completed">
              <input type="checkbox" ${isCompleted ? 'checked' : ''} onchange="window.taskApp.toggleTaskStatus('${task.id}')">
              <span class="task-item-title">${this.escapeHTML(task.title)}</span>
            </label>
            <div class="badges-group">
              <span class="badge ${priorityClass}">${task.priority}</span>
              <span class="badge ${statusClass}">${statusLabel}</span>
            </div>
          </div>
          
          ${task.description ? `<p class="task-item-desc">${this.escapeHTML(task.description)}</p>` : ''}

          <div class="task-item-footer">
            <div class="badges-group">
              <span class="badge badge-category">${task.category}</span>
              <span>Due: <strong>${task.dueDate}</strong></span>
              <span>Est: <strong>${task.estimatedHours}h</strong></span>
            </div>
            <div>
              <button class="btn btn-secondary btn-sm" onclick="window.taskApp.toggleTaskStatus('${task.id}')" title="Cycle status">
                Cycle Status
              </button>
              <button class="btn btn-danger-outline btn-sm" onclick="window.taskApp.deleteTask('${task.id}')" title="Delete task">
                Delete
              </button>
            </div>
          </div>
        </article>
      `;
    }).join("");
  }

  escapeHTML(str) {
    if (!str) return "";
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  showToast(message) {
    if (!this.toastBanner) return;
    this.toastBanner.textContent = message;
    this.toastBanner.classList.add("show");
    setTimeout(() => {
      this.toastBanner.classList.remove("show");
    }, 2800);
  }

  render() {
    this.renderStats();
    this.renderTaskList();
  }
}

// Instantiate on DOMContentLoaded
document.addEventListener("DOMContentLoaded", () => {
  window.taskApp = new TaskGravityApp();
});
