using NAudio.CoreAudioApi;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace AudioAttenuationApp.Helper
{
    public class VolumeFader : IDisposable
    {
        private float _currentVolume;
        private bool disposedValue;
        private readonly object _lock = new();
        private readonly SimpleAudioVolume simpleAudioVolume;

        public float CurrentVolume
        {
            get
            {
                lock (_lock) { return _currentVolume; }
            }
            private set
            {
                lock (_lock) { _currentVolume = Math.Clamp(value, 0.0f, 1.0f); }
            }
        }

        public VolumeFader(SimpleAudioVolume volume)
        {

            simpleAudioVolume = volume;
            CurrentVolume = volume.Volume;
        }

        /// <summary>
        /// Smoothly changes the volume to the target value over the specified duration.
        /// </summary>
        /// <param name="targetVolume">Target volume (0.0 to 1.0)</param>
        /// <param name="durationMs">Duration in milliseconds for the volume change</param>
        /// <param name="stepTimeMs">Time between volume updates in milliseconds</param>
        public async Task FadeToAsync(float targetVolume, int durationMs = 1000, int stepTimeMs = 50)
        {
            targetVolume = Math.Clamp(targetVolume, 0.0f, 1.0f);
            float startVolume = CurrentVolume;
            float diff = targetVolume - startVolume;
            int steps = durationMs / stepTimeMs;

            for (int i = 1; i <= steps; i++)
            {
                ObjectDisposedException.ThrowIf(disposedValue, this);
                float progress = (float)i / steps;
                CurrentVolume = startVolume + diff * progress;
                //Debug.WriteLine($"Volume: {CurrentVolume:0.00}");
                simpleAudioVolume.Volume = CurrentVolume;
                await Task.Delay(stepTimeMs);
            }

            CurrentVolume = targetVolume;
            //Debug.WriteLine($"Final Volume: {CurrentVolume:0.00}");
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }
                disposedValue = true;
            }
        }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}