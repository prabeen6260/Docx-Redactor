import { useState } from 'react';
import './App.css';

function App() {
  const [status, setStatus] = useState<{ type: 'idle' | 'loading' | 'success' | 'error'; message: string }>({
    type: 'idle',
    message: '',
  });

  const handleRedact = async () => {
    setStatus({ type: 'loading', message: 'Processing document...' });

    try {
      await Word.run(async (context) => {
        // Track Changes
        context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        
        // Adding Confidential Header
        const body = context.document.body;
        const headerText = 'CONFIDENTIAL DOCUMENT';
        let headerAdded = false;
        
        const firstParagraph = body.paragraphs.getFirst();
        firstParagraph.load('text');
        await context.sync();
        
        if (!firstParagraph.text.includes(headerText)) {
            const headerParagraph = body.insertParagraph(headerText, 'Start');
            headerParagraph.font.bold = true;
            headerParagraph.font.color = 'red';
            headerParagraph.font.size = 14;
            headerParagraph.alignment = Word.Alignment.centered;
            headerAdded = true;
            await context.sync();
        }

        // Redacting Sensitive Information
        const paragraphs = body.paragraphs;
        paragraphs.load('text');
        await context.sync();

        const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
        const ssnRegex = /\b\d{3}-\d{2}-\d{4}\b/g;
        const phoneRegex = /\b(?:\+?1[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}\b/g;
        const creditCardRegex = /\b\d{4}[-\s]?\d{4}[-\s]?\d{4}[-\s]?\d{4}\b/g;
        const dobRegex = /\b(0[1-9]|1[0-2])[\/\-](0[1-9]|[12]\d|3[01])[\/\-](19|20)\d{2}\b/g;
        const idPatternRegex = /\b(EMP|MRN|INS|ID)[-]?[\d]+([-][\d]+)*\b/gi;

        let redactedCount = 0;

        for (const paragraph of paragraphs.items) {
            const text = paragraph.text;
            const sensitiveData: string[] = [];

            let match;
            while ((match = emailRegex.exec(text)) !== null) {
                if (!sensitiveData.includes(match[0])) sensitiveData.push(match[0]);
            }
            while ((match = ssnRegex.exec(text)) !== null) {
                if (!sensitiveData.includes(match[0])) sensitiveData.push(match[0]);
            }
            while ((match = phoneRegex.exec(text)) !== null) {
                if (!sensitiveData.includes(match[0])) sensitiveData.push(match[0]);
            }
            while ((match = creditCardRegex.exec(text)) !== null) {
                if (!sensitiveData.includes(match[0])) sensitiveData.push(match[0]);
            }
            while ((match = dobRegex.exec(text)) !== null) {
                if (!sensitiveData.includes(match[0])) sensitiveData.push(match[0]);
            }
            while ((match = idPatternRegex.exec(text)) !== null) {
                if (!sensitiveData.includes(match[0])) sensitiveData.push(match[0]);
            }
            
            emailRegex.lastIndex = 0;
            ssnRegex.lastIndex = 0;
            phoneRegex.lastIndex = 0;
            creditCardRegex.lastIndex = 0;
            dobRegex.lastIndex = 0;
            idPatternRegex.lastIndex = 0;

            for (let i = 0; i < sensitiveData.length; i++) {
                const data = sensitiveData[i];
                const ranges = paragraph.search(data, { matchCase: false, matchWholeWord: false });
                ranges.load('items');
                await context.sync();

                for (const range of ranges.items) {
                    range.insertText('[REDACTED]', 'Replace');
                    redactedCount++;
                }
            }
        }

        await context.sync();
        
        const messages: string[] = [];
        if (redactedCount > 0) {
            messages.push(`Redacted ${redactedCount} sensitive item${redactedCount > 1 ? 's' : ''}`);
        } else {
            messages.push('No sensitive information found to redact');
        }
        if (headerAdded) {
            messages.push('added confidential header');
        } else {
            messages.push('header already present');
        }
        messages.push('Track Changes enabled');
        
        setStatus({ 
            type: 'success', 
            message: messages.join(', ') + '.'
        });
      });
    } catch (error: any) {
      console.error(error);
      if (error instanceof OfficeExtension.Error) {
        console.error('Debug info:', JSON.stringify(error.debugInfo));
      }
      setStatus({ type: 'error', message: `Error: ${error.message}` });
    }
  };

  return (
    <div className="container">
      <div className="card">
        <div className="header">
          <div className="badge">OFFICE ADD-IN</div>
          <h1 className="title">McCarren Redactor</h1>
          <p className="description">
            Automatically redact sensitive information (Emails, SSNs, Phone Numbers, Credit Cards, DOB, IDs) and mark the document as confidential.
          </p>
        </div>

        <div className="actions">
          <button 
            className="btn" 
            onClick={handleRedact} 
            disabled={status.type === 'loading'}
          >
            {status.type === 'loading' ? 'Processing...' : 'Redact & Protect'}
          </button>
        </div>

        {status.type !== 'idle' && (
          <div className={`status status-${status.type}`}>
            {status.message}
          </div>
        )}
      </div>
    </div>
  );
}

export default App;
