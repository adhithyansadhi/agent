document.addEventListener("DOMContentLoaded", () => {
  const form = document.getElementById("comparison-form");
  const compareButton = document.getElementById("compare-button");
  const buttonText = document.querySelector(".button-text");
  const spinner = document.querySelector(".spinner");
  const resultsSection = document.getElementById("results-section");
  const logWindow = document.getElementById("log-window");
  const downloadLink = document.getElementById("download-link");

  let eventSource;

  form.addEventListener("submit", async (e) => {
    e.preventDefault();

    // UI updates for processing
    buttonText.textContent = "Comparing...";
    spinner.style.display = "block";
    compareButton.disabled = true;
    resultsSection.style.display = "block";
    logWindow.textContent = "Initializing comparison...";
    downloadLink.style.display = "none";

    if (eventSource) {
      eventSource.close();
    }

    const formData = new FormData(form);

    try {
      const response = await fetch("/compare", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || "Failed to start comparison.");
      }

      const data = await response.json();
      const jobId = data.job_id;

      // Connect to the log streaming endpoint
      eventSource = new EventSource(`/stream/${jobId}`);

      eventSource.onmessage = (event) => {
        const message = event.data;

        if (message.startsWith(`DONE:`)) {
          const parts = message.split(":"); // e.g., ["DONE", "completed", "job-id"]
          const status = parts[1];

          if (status === "completed") {
            logWindow.textContent +=
              "\n\nProcess complete. Your download is ready.";
            downloadLink.href = `/download/${jobId}`;
            downloadLink.style.display = "inline-block";
          } else {
            // Handles 'failed' or any other status
            logWindow.textContent +=
              "\n\nProcess finished with an error. No download available.";
          }
          eventSource.close();
          resetButton();
        } else {
          if (logWindow.textContent === "Initializing comparison...") {
            logWindow.textContent = message;
          } else {
            logWindow.textContent += `\n${message}`;
          }
          // Auto-scroll to the bottom
          logWindow.scrollTop = logWindow.scrollHeight;
        }
      };

      // ### THIS IS THE DEBUGGING FIX ###
      eventSource.onerror = () => {
        logWindow.textContent +=
          "\n\n[DEVELOPER] Error: Connection to server lost. This can happen if the Flask server is restarted or crashes unexpectedly. Check the server console for errors.";
        eventSource.close();
        resetButton();
      };
      // ### END OF FIX ###
    } catch (error) {
      logWindow.textContent = `Error: ${error.message}`;
      resetButton();
    }
  });

  function resetButton() {
    buttonText.textContent = "Compare Folders";
    spinner.style.display = "none";
    compareButton.disabled = false;
  }
});
