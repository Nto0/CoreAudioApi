/*
  LICENSE
  -------
  Copyright (c) 2016-2018 Francois Gendron <fg@frgn.ca>
  Copyright (c) 2021 Tim Hoppmann <bahz@bahz.eu>

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

namespace CoreAudioApi
{

    // Class to interact with a MMDevice as an object with attributes
    public class AudioDevice
    {
        // Type is either "Playback" or "Recording"
        public AudioDeviceType Type { get; }
        // Name of the MMDevice ex: "Speakers (Realtek High Definition Audio)"
        public string Name { get; }
        // ID of the MMDevice ex: "{0.0.0.00000000}.{c4aadd95-74c7-4b3b-9508-b0ef36ff71ba}"
        public string ID { get; }
        // The MMDevice itself
        public MMDevice Device { get; }


        // To be created, a new AudioDevice needs an Index, and the MMDevice it will communicate with
        public AudioDevice(MMDevice baseDevice)
        {
            if (baseDevice == null)
                throw new ArgumentNullException(nameof(baseDevice));

            // If the received MMDevice is a playback device
            if (baseDevice.DataFlow == EDataFlow.eRender)
            {
                // Set this object's Type to "Playback"
                this.Type = AudioDeviceType.Playback;
            }
            // If not, if the received MMDevice is a recording device
            else if (baseDevice.DataFlow == EDataFlow.eCapture)
            {
                // Set this object's Type to "Recording"
                this.Type = AudioDeviceType.Recording;
            }

            // Set this object's Name to that of the received MMDevice's FriendlyName
            this.Name = baseDevice.FriendlyName;

            // Set this object's Device to the received MMDevice
            this.Device = baseDevice;

            // Set this object's ID to that of the received MMDevice's ID
            this.ID = baseDevice.ID;
        }

        public AudioDeviceUsage IsDefault()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            return GetDefaultDeviceUsage(ID, DevEnum, Type);
        }
        public bool IsDefault(AudioDeviceUsage usage)
        {
            return IsDefault().HasFlag(usage);
        }

        static AudioDeviceUsage GetDefaultDeviceUsage(string id, MMDeviceEnumerator devEnum, AudioDeviceType type)
        {
            if (type == AudioDeviceType.Playback)
                return GetDefaultDeviceUsage(id, devEnum, EDataFlow.eRender);
            else if (type == AudioDeviceType.Recording)
                return GetDefaultDeviceUsage(id, devEnum, EDataFlow.eCapture);
            return AudioDeviceUsage.None;
        }
        static AudioDeviceUsage GetDefaultDeviceUsage(string id, MMDeviceEnumerator devEnum, EDataFlow type)
        {
            string defaultSystemDeviceId = devEnum.GetDefaultAudioEndpoint(type, ERole.eConsole)?.ID;
            string defaultMediaDeviceId = devEnum.GetDefaultAudioEndpoint(type, ERole.eMultimedia)?.ID;
            string defaultComDeviceId = devEnum.GetDefaultAudioEndpoint(type, ERole.eCommunications)?.ID;
            return GetDefaultDeviceUsage(id, defaultSystemDeviceId, defaultMediaDeviceId, defaultComDeviceId);
        }
        static AudioDeviceUsage GetDefaultDeviceUsage(string id, string defaultSystemDeviceId, string defaultMediaDeviceId, string defaultComDeviceId)
        {
            var res = AudioDeviceUsage.None;
            if (IsIdEqual(id, defaultSystemDeviceId))
                res |= AudioDeviceUsage.System;
            if (IsIdEqual(id, defaultMediaDeviceId))
                res |= AudioDeviceUsage.Multimedia;
            if (IsIdEqual(id, defaultComDeviceId))
                res |= AudioDeviceUsage.Communication;
            return res;
        }

        public static List<AudioDevice> GetPlaybackDevices()
        {
            // Create a new MMDeviceEnumerator
            MMDeviceEnumerator devEnum = new MMDeviceEnumerator();
            // Create a MMDeviceCollection of every devices that are enabled
            MMDeviceCollection deviceCollection = devEnum.EnumerateAudioEndPoints(EDataFlow.eRender, EDeviceState.DEVICE_STATE_ACTIVE);
            // For every MMDevice in DeviceCollection
            var res = new List<AudioDevice>(deviceCollection.Count);
            for (int i = 0; i < deviceCollection.Count; i++)
            {
                var item = deviceCollection[i];
                // Output the result of the creation of a new AudioDevice while assining it an index, and the MMDevice itself, and the default value
                res.Add(new AudioDevice(item));
            }
            return res;
        }
        public static List<AudioDevice> GetRecordingDevices()
        {
            // Create a new MMDeviceEnumerator
            MMDeviceEnumerator devEnum = new MMDeviceEnumerator();
            // Create a MMDeviceCollection of every devices that are enabled
            MMDeviceCollection deviceCollection = devEnum.EnumerateAudioEndPoints(EDataFlow.eCapture, EDeviceState.DEVICE_STATE_ACTIVE);
            // For every MMDevice in DeviceCollection
            var res = new List<AudioDevice>(deviceCollection.Count);
            for (int i = 0; i < deviceCollection.Count; i++)
            {
                var item = deviceCollection[i];
                // Output the result of the creation of a new AudioDevice while assining it an index, and the MMDevice itself, and the default value
                res.Add(new AudioDevice(item));
            }
            return res;
        }
        public static AudioDevice GetDeviceById(string id)
        {
            // Create a new MMDeviceEnumerator
            MMDeviceEnumerator devEnum = new MMDeviceEnumerator();
            // Create a MMDeviceCollection of every devices that are enabled
            MMDeviceCollection deviceCollection = devEnum.EnumerateAudioEndPoints(EDataFlow.eAll, EDeviceState.DEVICE_STATE_ACTIVE);
            // For every MMDevice in DeviceCollection
            var res = new List<AudioDevice>(deviceCollection.Count);
            for (int i = 0; i < deviceCollection.Count; i++)
            {
                // If this MMDevice's ID is the same as the string received by the ID parameter
                if (IsIdEqual(deviceCollection[i].ID, id))
                {
                    var item = deviceCollection[i];
                    return new AudioDevice(item);
                }
            }
            return null;
        }

        static bool IsIdEqual(string a, string b)
        {
            if (string.IsNullOrEmpty(a))
                return string.IsNullOrEmpty(b);
            if (string.IsNullOrEmpty(b))
                return false;
            return (string.Compare(a, b, System.StringComparison.OrdinalIgnoreCase) == 0);
        }

        public static AudioDevice GetDefaultPlaybackDevice(AudioDeviceUsage usage = AudioDeviceUsage.System)
        {
            // Create a new MMDeviceEnumerator
            MMDeviceEnumerator devEnum = new MMDeviceEnumerator();
            // Create a MMDeviceCollection of every devices that are enabled
            MMDeviceCollection deviceCollection = devEnum.EnumerateAudioEndPoints(EDataFlow.eRender, EDeviceState.DEVICE_STATE_ACTIVE);
            // If this MMDevice's ID is either, the same the default playback device's ID, or the same as the default recording device's ID
            GetDefaultDeviceIds(devEnum, usage, EDataFlow.eRender, out string defaultSystemDeviceId, out string defaultMediaDeviceId, out string defaultCommDeviceId);
            // For every MMDevice in DeviceCollection
            for (int i = 0; i < deviceCollection.Count; i++)
            {
                // Output the result of the creation of a new AudioDevice while assining it an index, and the MMDevice itself, and the default value
                if (MatchAnyId(deviceCollection[i].ID, defaultSystemDeviceId, defaultMediaDeviceId, defaultCommDeviceId))
                    return new AudioDevice(deviceCollection[i]);
            }
            return null;
        }

        public static AudioDevice GetDefaultRecordingDevice(AudioDeviceUsage usage = AudioDeviceUsage.System)
        {
            if (usage == AudioDeviceUsage.None)
                return null;
            // Create a new MMDeviceEnumerator
            MMDeviceEnumerator devEnum = new MMDeviceEnumerator();
            // Create a MMDeviceCollection of every devices that are enabled
            MMDeviceCollection deviceCollection = devEnum.EnumerateAudioEndPoints(EDataFlow.eCapture, EDeviceState.DEVICE_STATE_ACTIVE);
            // If this MMDevice's ID is either, the same the default playback device's ID, or the same as the default recording device's ID
            GetDefaultDeviceIds(devEnum, usage, EDataFlow.eCapture, out string defaultSystemDeviceId, out string defaultMediaDeviceId, out string defaultCommDeviceId);
            // For every MMDevice in DeviceCollection
            for (int i = 0; i < deviceCollection.Count; i++)
            {
                // Output the result of the creation of a new AudioDevice while assining it an index, and the MMDevice itself, and the default value
                if (MatchAnyId(deviceCollection[i].ID, defaultSystemDeviceId, defaultMediaDeviceId, defaultCommDeviceId))
                    return new AudioDevice(deviceCollection[i]);
            }
            return null;
        }

        static bool MatchAnyId(string id, string comp1, string comp2, string comp3)
        {
            return IsIdEqual(id, comp1) || IsIdEqual(id, comp2) || IsIdEqual(id, comp3);
        }

        static void GetDefaultDeviceIds(MMDeviceEnumerator devEnum, AudioDeviceUsage usage, EDataFlow flow, out string systemDeviceId, out string mediaDeviceId, out string commDeviceId)
        {
            systemDeviceId = ((usage & AudioDeviceUsage.System) != 0 ? devEnum.GetDefaultAudioEndpoint(flow, ERole.eConsole)?.ID : null);
            mediaDeviceId = ((usage & AudioDeviceUsage.Multimedia) != 0 ? devEnum.GetDefaultAudioEndpoint(flow, ERole.eMultimedia)?.ID : null);
            commDeviceId = ((usage & AudioDeviceUsage.Communication) != 0 ? devEnum.GetDefaultAudioEndpoint(flow, ERole.eCommunications)?.ID : null);
        }


        public static bool SetDefaultPlaybackDevice(string deviceId, AudioDeviceUsage usage = AudioDeviceUsage.All)
        {
            if (string.IsNullOrEmpty(deviceId))
                throw new ArgumentNullException();
            if (usage == AudioDeviceUsage.None)
                return false;
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            // Create a MMDeviceCollection of every devices that are enabled
            MMDeviceCollection DeviceCollection = DevEnum.EnumerateAudioEndPoints(EDataFlow.eRender, EDeviceState.DEVICE_STATE_ACTIVE);
            for (int i = 0; i < DeviceCollection.Count; i++)
            {
                // If this MMDevice's ID is the same as the ID of the MMDevice received by the InputObject parameter
                if (DeviceCollection[i].ID == deviceId)
                {
                    // Create a new audio PolicyConfigClient
                    PolicyConfigClient client = new PolicyConfigClient();
                    // Using PolicyConfigClient, set the given device as the default playback device
                    if ((usage & AudioDeviceUsage.System) == AudioDeviceUsage.System)
                        client.SetDefaultEndpoint(DeviceCollection[i].ID, ERole.eConsole);
                    if ((usage & AudioDeviceUsage.Multimedia) == AudioDeviceUsage.Multimedia)
                        client.SetDefaultEndpoint(DeviceCollection[i].ID, ERole.eMultimedia);
                    if ((usage & AudioDeviceUsage.Communication) == AudioDeviceUsage.Communication)
                        client.SetDefaultEndpoint(DeviceCollection[i].ID, ERole.eCommunications);
                    return true;
                }
            }
            return false;
        }

        public static bool SetDefaultRecordingDevice(string deviceId, AudioDeviceUsage usage = AudioDeviceUsage.All)
        {
            if (string.IsNullOrEmpty(deviceId))
                throw new ArgumentNullException();
            if (usage == AudioDeviceUsage.None)
                return false;
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            // Create a MMDeviceCollection of every devices that are enabled
            MMDeviceCollection DeviceCollection = DevEnum.EnumerateAudioEndPoints(EDataFlow.eCapture, EDeviceState.DEVICE_STATE_ACTIVE);
            for (int i = 0; i < DeviceCollection.Count; i++)
            {
                // If this MMDevice's ID is the same as the ID of the MMDevice received by the InputObject parameter
                if (DeviceCollection[i].ID == deviceId)
                {
                    // Create a new audio PolicyConfigClient
                    PolicyConfigClient client = new PolicyConfigClient();
                    // Using PolicyConfigClient, set the given device as the default playback device
                    if ((usage & AudioDeviceUsage.System) == AudioDeviceUsage.System)
                        client.SetDefaultEndpoint(DeviceCollection[i].ID, ERole.eConsole);
                    if ((usage & AudioDeviceUsage.Multimedia) == AudioDeviceUsage.Multimedia)
                        client.SetDefaultEndpoint(DeviceCollection[i].ID, ERole.eMultimedia);
                    if ((usage & AudioDeviceUsage.Communication) == AudioDeviceUsage.Communication)
                        client.SetDefaultEndpoint(DeviceCollection[i].ID, ERole.eCommunications);
                    return true;
                }
            }
            return false;
        }

    }

}
