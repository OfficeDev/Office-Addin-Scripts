import * as appInsights from 'applicationinsights';
const noElapsedTime = 0;
try {
    let insight = appInsights.getClient('de0d9e7c-1f46-4552-bc21-4e43e489a015');
insight.trackEvent('Host', { Host: 'PowerPoint' }, { noElapsedTime });
} catch (err) {
    throw new Error(err);
}


