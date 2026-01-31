// Auto-scroll behavior for the chat messages list
// Ensures users always see the latest messages without manual scrolling
console.log('ChatMessageList.razor.js loaded with improved auto-scroll!');

// Global scroll function that can be called from Blazor
window.scrollChatToBottom = function(smooth = true) {
    // Strategy 1: Scroll the message-list-container
    const scrollContainer = document.querySelector('.message-list-container');
    if (scrollContainer) {
        scrollContainer.scrollTo({
            top: scrollContainer.scrollHeight,
            behavior: smooth ? 'smooth' : 'instant'
        });
    }
    
    // Strategy 2: Scroll anchor element into view
    const scrollAnchor = document.querySelector('.scroll-anchor');
    if (scrollAnchor) {
        scrollAnchor.scrollIntoView({ 
            behavior: smooth ? 'smooth' : 'instant',
            block: 'end'
        });
    }
    
    // Strategy 3: Scroll the window/body as fallback
    if (!scrollContainer) {
        window.scrollTo({
            top: document.body.scrollHeight,
            behavior: smooth ? 'smooth' : 'instant'
        });
    }
};

// Custom element for additional MutationObserver-based scrolling
window.customElements.define('chat-messages', class ChatMessages extends HTMLElement {
    static _isFirstAutoScroll = true;

    connectedCallback() {
        this._observer = new MutationObserver(mutations => this._scheduleAutoScroll(mutations));
        // Observe changes including subtree for streaming content
        this._observer.observe(this, { 
            childList: true, 
            attributes: true,
            subtree: true,
            characterData: true 
        });
        
        // Initial scroll
        setTimeout(() => window.scrollChatToBottom(false), 100);
    }

    disconnectedCallback() {
        this._observer.disconnect();
    }

    _scheduleAutoScroll(mutations) {
        // Debounce calls
        cancelAnimationFrame(this._nextAutoScroll);
        this._nextAutoScroll = requestAnimationFrame(() => {
            const hasRelevantChange = mutations.some(m => {
                // Text content changes (streaming)
                if (m.type === 'characterData') return true;
                
                // New nodes added
                if (m.type === 'childList' && m.addedNodes.length > 0) {
                    return Array.from(m.addedNodes).some(n => {
                        if (n.nodeType === Node.TEXT_NODE) return true;
                        if (n.nodeType !== Node.ELEMENT_NODE) return false;
                        
                        // Check for message-related classes
                        const messageClasses = [
                            'user-message', 'assistant-message', 'assistant-message-text',
                            'waiting-message', 'adaptive-card-container', 'loading-spinner'
                        ];
                        
                        if (n.classList && messageClasses.some(c => n.classList.contains(c))) {
                            return true;
                        }
                        
                        // Check for nested message elements
                        if (n.querySelector) {
                            return messageClasses.some(c => n.querySelector('.' + c)) ||
                                   n.querySelector('loading-spinner');
                        }
                        
                        return false;
                    });
                }
                return false;
            });
            
            // Scroll if content changed or user is near bottom
            if (hasRelevantChange || this._isNearBottom(200)) {
                window.scrollChatToBottom(true);
            }
        });
    }

    _isNearBottom(threshold = 200) {
        const scrollContainer = document.querySelector('.message-list-container');
        if (scrollContainer) {
            const maxScroll = scrollContainer.scrollHeight - scrollContainer.clientHeight;
            const remaining = maxScroll - scrollContainer.scrollTop;
            return remaining < threshold;
        }
        // Fallback: check window scroll
        const maxScroll = document.body.scrollHeight - window.innerHeight;
        const remaining = maxScroll - window.scrollY;
        return remaining < threshold;
    }
});

// Also scroll when the page loads
document.addEventListener('DOMContentLoaded', () => {
    setTimeout(() => window.scrollChatToBottom(false), 200);
});
