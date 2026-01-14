// config.gs

// --- Global Configuration ---
function getConfig() {
  return {
    EXPORT_FOLDER_NAME: "exportemails",
    FILENAME_SUFFIX: "_export",
    MAX_RETRIES: 3,
    RETRY_DELAY: 2000,
    BATCH_SIZE: 50,
    INITIAL_BATCH_SIZE: 10,
    MAX_CHARS_ALLOWED: 900000,
    SLEEP_BETWEEN_BATCHES: 800,
    MAX_MESSAGE_LENGTH: 25000,

    // --- Label and Document Configuration ---
    LABEL_CONFIG: [
      {
        label: "customer/sample_customer",
        docBaseName: "Sample_Customer_Mail"
      },

      // Add configurations for other labels here
    ]
  };
}
