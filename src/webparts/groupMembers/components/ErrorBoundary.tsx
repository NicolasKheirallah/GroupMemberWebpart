import * as React from 'react';
import { MessageBar, MessageBarType, DefaultButton } from '@fluentui/react';

interface ErrorBoundaryProps {
  children?: React.ReactNode;
}

interface ErrorBoundaryState {
  hasError: boolean;
  error: Error | undefined;
  errorInfo: React.ErrorInfo | undefined;
}

export class ErrorBoundary extends React.Component<ErrorBoundaryProps, ErrorBoundaryState> {
  constructor(props: ErrorBoundaryProps) {
    super(props);
    this.state = {
      hasError: false,
      error: undefined,
      errorInfo: undefined
    };
  }

  static getDerivedStateFromError(error: Error): Partial<ErrorBoundaryState> {
    return {
      hasError: true,
      error
    };
  }

  public componentDidCatch(error: Error, errorInfo: React.ErrorInfo): void {
    this.setState({
      error,
      errorInfo
    });

    // Log error for debugging
    console.error('ErrorBoundary caught an error:', error, errorInfo);
  }

  private handleRetry = (): void => {
    this.setState({
      hasError: false,
      error: undefined,
      errorInfo: undefined
    });
  };

  public render(): React.ReactNode {
    if (this.state.hasError) {
      return (
        <div style={{ padding: '20px' }}>
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={true}
            actions={
              <div>
                <DefaultButton
                  onClick={this.handleRetry}
                  text="Try Again"
                />
              </div>
            }
          >
            <strong>Something went wrong while loading the Group Members web part.</strong>
            <br />
            {this.state.error?.message && (
              <>
                <br />
                <details>
                  <summary>Error details (for administrators)</summary>
                  <pre style={{ 
                    fontSize: '12px', 
                    overflow: 'auto', 
                    maxHeight: '200px',
                    backgroundColor: '#f5f5f5',
                    padding: '10px',
                    marginTop: '10px'
                  }}>
                    {this.state.error.message}
                    {this.state.errorInfo?.componentStack}
                  </pre>
                </details>
              </>
            )}
          </MessageBar>
        </div>
      );
    }

    return this.props.children;
  }
}