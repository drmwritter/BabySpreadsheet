import React from 'react';
import Box from '@mui/material/Box';
import Typography from '@mui/material/Typography';
import Button from '@mui/material/Button';

class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, error: null, errorInfo: null };
  }

  static getDerivedStateFromError(error) {
    return { hasError: true, error };
  }

  componentDidCatch(error, errorInfo) {
    this.setState({ errorInfo });
    // You can also log the error to an error reporting service
    console.error("Uncaught error:", error, errorInfo);
  }

  render() {
    if (this.state.hasError) {
      return (
        <Box
          sx={{
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            padding: 4,
            textAlign: 'center',
            backgroundColor: '#ffebee',
            color: '#c62828',
            borderRadius: 2,
            border: '1px solid #ef9a9a',
            margin: 2,
          }}
        >
          <Typography variant="h5" gutterBottom>
            Oops! Something went wrong.
          </Typography>
          <Typography variant="body1">
            An unexpected error occurred. You can try reloading the page.
          </Typography>
          {this.state.error && (
            <pre style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-all', textAlign: 'left', backgroundColor: '#ffcdd2', padding: '10px', borderRadius: '4px', margin: '20px 0', maxWidth: '100%' }}>
              {this.state.error.toString()}
              <br />
              {this.state.errorInfo?.componentStack}
            </pre>
          )}
          <Button variant="contained" color="error" onClick={() => window.location.reload()}>
            Reload Page
          </Button>
        </Box>
      );
    }

    return this.props.children;
  }
}

export default ErrorBoundary;
