/**
 * Welcome Notification System
 * Manages dismissible, one-time welcome notifications for version updates and introductions
 */

class WelcomeNotification {
  constructor() {
    this.currentVersion = chrome.runtime.getManifest().version;
    this.storageKey = 'welcomeNotifications';
    this.init();
  }

  async init() {
    const shouldShow = await this.shouldShowWelcome();
    if (shouldShow) {
      this.showWelcomeModal();
    }
  }

  async shouldShowWelcome() {
    return new Promise((resolve) => {
      chrome.storage.local.get([this.storageKey], (result) => {
        const notifications = result[this.storageKey] || {};
        const lastSeenVersion = notifications.lastSeenVersion;
        
        // Show if first time user or version has changed
        const shouldShow = !lastSeenVersion || lastSeenVersion !== this.currentVersion;
        resolve(shouldShow);
      });
    });
  }

  async markAsSeen() {
    return new Promise((resolve) => {
      chrome.storage.local.get([this.storageKey], (result) => {
        const notifications = result[this.storageKey] || {};
        notifications.lastSeenVersion = this.currentVersion;
        notifications.dismissedAt = new Date().toISOString();
        
        chrome.storage.local.set({ [this.storageKey]: notifications }, () => {
          resolve();
        });
      });
    });
  }

  getWelcomeContent() {
    // Use external content configuration if available, otherwise fallback to inline
    const contentSource = window.WELCOME_CONTENT || this.getInlineContent();
    const content = contentSource[this.currentVersion] || contentSource['default'];
    
    return content;
  }

  getInlineContent() {
    // Fallback content if external file isn't loaded
    return {
      'default': {
        title: '🎉 Welcome to Intune Admin Center Tools!',
        intro: `Thank you for using Intune Admin Center Tools! This extension helps you efficiently manage Intune devices and assignments.`,
        changelog: [
          '🔧 General improvements and bug fixes'
        ],
        features: [
          '🔍 Search and manage Azure AD groups',
          '📱 Check device/user assignments',
          '👥 Manage group memberships',
          '🎨 Dark/Light theme support'
        ],
        tips: [
          'Use Ctrl+Shift+W to show this welcome message anytime',
          'The extension works on any Intune Admin Center page'
        ]
      }
    };
  }

  showWelcomeModal() {
    const content = this.getWelcomeContent();
    
    // Create modal HTML
    const modalHTML = `
      <div id="welcomeModal" class="welcome-modal-overlay">
        <div class="welcome-modal">
          <div class="welcome-header">
            <h4>${content.title}</h4>
            <button id="closeWelcome" class="close-btn">
              <i class="material-icons">close</i>
            </button>
          </div>
          
          <div class="welcome-content">
            <div class="welcome-section">
              <p class="welcome-intro">${content.intro}</p>
            </div>
            
            <div class="welcome-section">
              
            
            ${content.changelog && content.changelog.length > 0 ? `
            <div class="welcome-section">
              <h5>📋 What's New in v${this.currentVersion}:</h5>
              <ul class="changelog-list">
                ${content.changelog.map(item => `<li>${item}</li>`).join('')}
              </ul>
            </div>
            ` : ''}
            <h5>🚀 Key Features:</h5>
              <ul class="feature-list">
                ${content.features.map(feature => `<li>${feature}</li>`).join('')}
              </ul>
            </div>
            ${content.tips && content.tips.length > 0 ? `
            <div class="welcome-section">
              <h5>💡 Pro Tips:</h5>
              <ul class="tips-list">
                ${content.tips.map(tip => `<li>${tip}</li>`).join('')}
              </ul>
            </div>
            ` : ''}
            
            <div class="welcome-section">
              <p class="welcome-tip">
                � <strong>Quick Start:</strong> This extension works directly within the Intune Admin Center. 
                Navigate to any device or user page to access these powerful management tools!
              </p>
            </div>
          </div>
          
          <div class="welcome-actions">
            <button id="dismissWelcome" class="btn waves-effect waves-light cyan lighten-1">
              Got it, thanks!
              <i class="material-icons right">check</i>
            </button>
          </div>
        </div>
      </div>
    `;

    // Add modal to DOM
    document.body.insertAdjacentHTML('beforeend', modalHTML);

    // Add event listeners
    this.addEventListeners();

    // Show modal with animation
    setTimeout(() => {
      document.getElementById('welcomeModal').classList.add('show');
    }, 10);
  }

  addEventListeners() {
    const modal = document.getElementById('welcomeModal');
    const closeBtn = document.getElementById('closeWelcome');
    const dismissBtn = document.getElementById('dismissWelcome');

    // Close button handler
    closeBtn.addEventListener('click', () => this.dismissModal());
    
    // Dismiss button handler
    dismissBtn.addEventListener('click', () => this.dismissModal());
    
    // Click outside to close
    modal.addEventListener('click', (e) => {
      if (e.target === modal) {
        this.dismissModal();
      }
    });

    // ESC key to close
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && document.getElementById('welcomeModal')) {
        this.closeModal();
      }
    });
  }

  closeModal() {
    const modal = document.getElementById('welcomeModal');
    if (modal) {
      modal.classList.remove('show');
      setTimeout(() => {
        modal.remove();
      }, 300);
    }
  }

  async dismissModal() {
    await this.markAsSeen();
    this.closeModal();
    
    // Show a brief confirmation
    if (typeof showNotification === 'function') {
      showNotification('Welcome message dismissed. You can always find help in the extension menu!', 'info');
    }
  }

  // Static method to manually show welcome (for testing or user request)
  static async showManual() {
    const welcome = new WelcomeNotification();
    welcome.showWelcomeModal();
  }

  // Static method to reset welcome status (for testing)
  static async reset() {
    chrome.storage.local.remove('welcomeNotifications', () => {
      console.log('Welcome notification status reset');
    });
  }
}

// Initialize when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    new WelcomeNotification();
  });
} else {
  new WelcomeNotification();
}

// Export for external use
window.WelcomeNotification = WelcomeNotification;
