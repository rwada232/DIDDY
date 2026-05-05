let compressor, filter, source;

function createEnhancedStream(stream, audioContext) {
    source = audioContext.createMediaStreamSource(stream);
    const destination = audioContext.createMediaStreamDestination();

    // Compressor (Fixes "Chopping")
    compressor = audioContext.createDynamicsCompressor();
    compressor.threshold.setValueAtTime(-24, audioContext.currentTime);
    compressor.ratio.setValueAtTime(12, audioContext.currentTime);

    // Filter (Noise Suppression)
    filter = audioContext.createBiquadFilter();
    filter.type = "highpass";
    filter.frequency.setValueAtTime(100, audioContext.currentTime);

    // Initial Chain
    source.connect(filter);
    filter.connect(compressor);
    compressor.connect(destination);

    return destination.stream;
}

// Settings Toggles
function updateAudioSettings(settings) {
    if (compressor) {
        compressor.threshold.setValueAtTime(settings.limiter ? -24 : -100, 0);
    }
    if (filter) {
        filter.frequency.setValueAtTime(settings.noiseSuppression ? 100 : 0, 0);
    }
}
