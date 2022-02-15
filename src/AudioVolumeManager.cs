/*
  LICENSE
  -------
  Copyright (c) 2021-2022 Tim Hoppmann <N@Nto0.net>

  Permission is hereby granted, free of charge, to any person obtaining a copy
  of this software and associated documentation files (the "Software"), to deal
  in the Software without restriction, including without limitation the rights
  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
  copies of the Software, and to permit persons to whom the Software is
  furnished to do so, subject to the following conditions:

  The above copyright notice and this permission notice shall be included in all
  copies or substantial portions of the Software.

  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
  SOFTWARE.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CoreAudioApi
{

    public static class AudioVolumeManager
    {

        /// <summary>
        /// Gets the current master volume in scalar values (percentage)
        /// </summary>
        /// <returns>NaN in case of an error, if successful the value will be between 0 and 1</returns>
        public static float GetMasterVolume()
        {
            var masterVol = GetMasterVolumeObject();
            if (masterVol == null)
                return float.NaN;
            return masterVol.MasterVolumeLevelScalar;
        }

        /// <summary>
        /// Gets the mute state of the master volume. 
        /// While the volume can be muted the <see cref="GetMasterVolume"/> will still return the pre-muted volume value.
        /// </summary>
        /// <returns>false if not muted, true if volume is muted</returns>
        public static bool GetMasterVolumeMute()
        {
            var masterVol = GetMasterVolumeObject();
            if (masterVol == null)
                return false;
            return masterVol.Mute;
        }

        /// <summary>
        /// Sets the master volume to a specific level
        /// </summary>
        /// <param name="volume">Value between 0 and 1 indicating the desired scalar value of the volume</param>
        public static void SetMasterVolume(float volume)
        {
            var masterVol = GetMasterVolumeObject();
            if (masterVol == null)
                return;
            masterVol.MasterVolumeLevelScalar = volume;
        }

        /// <summary>
        /// Increments or decrements the current volume level by the <see cref="amount"/>.
        /// </summary>
        /// <param name="amount">Value between -1 and 1 indicating the desired step amount. Use negative numbers to decrease
        /// the volume and positive numbers to increase it.</param>
        /// <returns>the new volume level assigned</returns>
        public static float ModifyMasterVolume(float amount)
        {
            var masterVol = GetMasterVolumeObject();
            if (masterVol == null)
                return float.NaN;

            // Get the level
            float volumeLevel = masterVol.MasterVolumeLevelScalar;

            // Calculate the new level
            float newLevel = volumeLevel + amount;
            newLevel = Math.Min(1, newLevel);
            newLevel = Math.Max(0, newLevel);

            masterVol.MasterVolumeLevelScalar = newLevel;

            // Return the new volume level that was set
            return newLevel;
        }

        /// <summary>
        /// Mute or unmute the master volume
        /// </summary>
        /// <param name="mute">true to mute the master volume, false to unmute</param>
        public static void SetMasterVolumeMute(bool mute)
        {
            var masterVol = GetMasterVolumeObject();
            if (masterVol == null)
                return;
            masterVol.Mute = mute;
        }

        /// <summary>
        /// Switches between the master volume mute states depending on the current state
        /// </summary>
        /// <returns>the current mute state, true if the volume was muted, false if unmuted</returns>
        public static bool ToggleMasterVolumeMute()
        {
            var masterVol = GetMasterVolumeObject();
            if (masterVol == null)
                return false;
            var mute = !masterVol.Mute;
            masterVol.Mute = mute;
            return mute;
        }

        public static AudioEndpointVolume GetMasterVolumeObject()
        {
            var deviceEnumerator = new MMDeviceEnumerator();
            var speakers = deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);
            return speakers?.AudioEndpointVolume;
        }


        // Individual Application Volume Manipulation

        public static float? GetApplicationVolume(int processId)
        {
            var volume = GetApplicationVolumeObject(processId);
            if (volume == null)
                return null;
            return volume.MasterVolume;
        }

        public static bool? GetApplicationMute(int processId)
        {
            var volume = GetApplicationVolumeObject(processId);
            if (volume == null)
                return null;
            return volume.Mute;
        }

        public static void SetApplicationVolume(int processId, float volume)
        {
            var obj = GetApplicationVolumeObject(processId);
            if (obj == null)
                return;
            obj.MasterVolume = volume;
        }

        public static void SetApplicationMute(int processId, bool mute)
        {
            var volume = GetApplicationVolumeObject(processId);
            if (volume == null)
                return;
            volume.Mute = mute;
        }

        public static bool ToggleApplicationMute(int processId)
        {
            var volume = GetApplicationVolumeObject(processId);
            if (volume == null)
                return false;
            var mute = !volume.Mute;
            volume.Mute = mute;
            return mute;
        }

        public static SimpleAudioVolume GetApplicationVolumeObject(int processId)
        {
            // get the speakers (1st render + multimedia) device
            var deviceEnumerator = new MMDeviceEnumerator();
            var speakers = deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);
            var sessions = speakers.AudioSessionManager.Sessions;

            // search for an audio session with the required process-id
            for (int i = 0; i < sessions.Count; ++i)
            {
                if (sessions[i].ProcessID == processId)
                    return sessions[i].SimpleAudioVolume;
            }
            return null;
        }

    }

}
