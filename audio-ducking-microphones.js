/**************************************************************
 * Audio Ducking Macro (Merged)
 *
 * Base macro:
 *   - Audio Ducking Macro by William Mills (Cisco)
 *   - Version 1-3-0 (Released: 03/02/26)
 *     - Added audio event timeouts to handle cases where monitored inputs
 *       are muted externally (no VU meter events) so unducking can still occur.
 *     - MTR call state support.
 *
 * This forked version keeps the upstream improvements and includes local
 * enhancements from "audio-ducking of microphones.js":
 * Version 0-0-1 (Released:10/03/26)
 *   - Optionally run VU monitoring in and out-of-call
 *   - Optional standby stop/restart
 *   - Added audienceOnly mode
 *   - Added configurable VU meter source (BeforeAEC/AfterAEC)
 *   - Optional NoiseRemoval.Mode config
 *   - Consolidated mic control with state caching to reduce repeated writes
 **************************************************************/

import xapi from 'xapi';

/**************************************************************
 * Configure the settings below
 **************************************************************/
const config = {
  button: { // Customise the macros control button name, color and icon
    name: 'Audio Modes', // Button/Panel/Alert Name
    color: '#f58142', // Button Color
    icon: 'Sliders', // Button Icon
    location: 'ControlPanel' // CallControls | ControlPanel | Hidden
  },

  showAlerts: true, // true = show alert messages on controller

  modeNames: {
    autoDuck: 'Automatisk',
    presentersOnly: 'Kun tale mikrofoner',
    audienceOnly: 'Kun tak mikrofoner',
    presentersAndAudience: 'Tale og tak mikrofoner'
  },

  defaultMode: 'autoDuck',

  // Only run macro when device is in a call
  OnlyRunWhenInCall: false,

  // Reset mode when a new call starts
  resetModeOnNewCall: true,

  // Stop monitoring in standby and restart on wake
  stopInStandby: false,

  // VU meter tap point (before/after echo cancellation)
  vuMeter: { source: 'BeforeAEC' }, // BeforeAEC | AfterAEC

  // Noise reduction / background noise removal
  noiseRemoval: { mode: 'Enabled' }, // Disabled | Enabled | Manual

  // Mics that are monitored for VU meter trigger levels
  // Example types: Microphone | Ethernet | USBMicrophone
  mics: [
    { ConnectorType: 'Microphone', ConnectorId: 1 },
    { ConnectorType: 'Microphone', ConnectorId: 2 },
    { ConnectorType: 'Microphone', ConnectorId: 3 }
  ],

  // Mics that are ducked/unducked (controlled)
  duck: [
    { ConnectorType: 'Microphone', ConnectorId: 1 },
    { ConnectorType: 'Microphone', ConnectorId: 2 },
    { ConnectorType: 'Microphone', ConnectorId: 3 }
  ],

  // Thresholds in which the monitored mic is considered high or low
  threshold: { high: 20, low: 5 },

  // Gain/Levels which should be set ducked or unducked
  levels: { duck: 0, unduck: 60 },

  // Duration where the monitored mic is low before unducking
  unduck: { timeout: 0 },

  // Samples taken every 100ms, 4 samples at 100ms = 400ms
  samples: 4,

  debug: false,

  panelId: 'audioDucking'
};

/**************************************************************
 * Do not change below
 **************************************************************/
const startVuMeterConnectors = consolidateConnectors(config.mics);
const micNames = createMicStrings(config.mics);

const sampleInterval = 100;
const averageLogFequency = 20;
const maxMissingAudioEvents = 100;

let gainLevel = 'Gain';
let unduckTimer;
let audioEventTimeout;
let missingAudioEventCounter = 0;
let listener;
let micLevels;
let micLevelAverages;
let audioEventCount = 0;
let mode;
let callId;

// Default VU meter source
let vuMeterSource = 'BeforeAEC';

// Cache to reduce repeated config writes
const groupState = {
  monitoredMics: null,  // 'duck' | 'unduck'
  controlledMics: null  // 'duck' | 'unduck'
};

setTimeout(init, 3000);

async function init() {
  gainLevel = await checkGainLevel();
  applyVuMeterConfig();
  await applyNoiseRemovalConfig();
  await createPanel();

  xapi.Event.UserInterface.Extensions.Widget.Action.on(processActions);

  // RoomOS call state
  xapi.Status.Call.on(({ ghost, Status, id }) => {
    if (Status && Status === 'Connected' && callId !== id) {
      callId = id;
      processNewCall('RoomOS');
      return;
    }
    if (ghost) processCallEnd('RoomOS');
  });

  // MTR call state
  xapi.Status.MicrosoftTeams.Calling.InCall.on((inMTRCall) => {
    if (inMTRCall === 'True') processNewCall('MTR');
    else processCallEnd('MTR');
  });

  // Standby handling
  if (config.stopInStandby) {
    xapi.Status.Standby.State.on((state) => {
      const isStandby = (String(state).toLowerCase() !== 'off');
      if (isStandby) {
        if (config.debug) console.log('Standby entered -> stopping monitor');
        stopMonitor();
      } else {
        if (config.debug) console.log('Wake detected -> applying mode');
        applyMode();
      }
    });
  }

  // Restore last selected mode (or default)
  const widgets = await xapi.Status.UserInterface.Extensions.Widget.get();
  const selection = widgets.find(w => w.WidgetId === config.panelId);
  const value = selection?.Value;
  mode = value && value !== '' ? value : config.defaultMode;

  await xapi.Command.UserInterface.Extensions.Widget.SetValue({
    WidgetId: config.panelId,
    Value: mode
  });

  applyMode();
}

function applyVuMeterConfig() {
  const allowedSources = ['BeforeAEC', 'AfterAEC'];
  const desired = config.vuMeter?.source ?? 'BeforeAEC';
  vuMeterSource = allowedSources.includes(desired) ? desired : 'BeforeAEC';

  if (!allowedSources.includes(desired) && config.debug) {
    console.warn(`Invalid vuMeter.source="${desired}". Falling back to "BeforeAEC".`);
  }
  if (config.debug) console.log(`VU meter Source=${vuMeterSource}, IntervalMs=${sampleInterval}`);
}

async function applyNoiseRemovalConfig() {
  const allowedModes = ['Disabled', 'Enabled', 'Manual'];
  const desired = config.noiseRemoval?.mode ?? 'Enabled';
  const modeToApply = allowedModes.includes(desired) ? desired : 'Enabled';

  try {
    await xapi.Config.Audio.Microphones.NoiseRemoval.Mode.set(modeToApply);
    if (config.debug) console.log(`NoiseRemoval.Mode set to: ${modeToApply}`);
  } catch (e) {
    if (config.debug) console.warn('Noise removal could not be set (may be unsupported):', e?.message ?? e);
  }
}

async function applyMode() {
  if (config.OnlyRunWhenInCall) {
    const inCall = await checkInCall();
    if (!inCall) return;
  }

  if (config.debug) console.log('Applying Mode:', mode);

  if (mode === 'presentersOnly') {
    stopMonitor();
    await setMicsLevel('monitored', 'unduck', true);
    await setMicsLevel('controlled', 'duck', true);
    return;
  }

  if (mode === 'presentersAndAudience') {
    stopMonitor();
    await setMicsLevel('monitored', 'unduck', true);
    await setMicsLevel('controlled', 'unduck', true);
    return;
  }

  if (mode === 'audienceOnly') {
    stopMonitor();
    await setMicsLevel('monitored', 'duck', true);
    await setMicsLevel('controlled', 'unduck', true);
    return;
  }

  // autoDuck
  await setMicsLevel('monitored', 'unduck', true);
  await setMicsLevel('controlled', 'unduck', true);
  startMonitor();
}

function processActions({ Type, Value, WidgetId }) {
  if (Type !== 'released') return;
  if (WidgetId !== config.panelId) return;
  if (mode === Value) return;

  mode = Value;
  applyMode();
}

async function checkInCall() {
  const mtrCall = await xapi.Status.MicrosoftTeams.Calling.InCall.get();
  const call = await xapi.Status.Call.get();
  return (call?.[0]?.Status === 'Connected') || (mtrCall === 'True');
}

function createMicLevels(samples) {
  const result = {};
  for (const key of micNames) result[key] = new Array(samples).fill(0);
  return result;
}

function processNewCall(callType) {
  if (config.debug) console.log(callType, 'Call Connected');

  if (config.resetModeOnNewCall) {
    mode = config.defaultMode;
    const modeName = config.modeNames[mode];
    const buttonName = config.button.name;

    xapi.Command.UserInterface.Extensions.Widget.SetValue({ WidgetId: config.panelId, Value: mode });
    applyMode();
    alert(`New Call Detected<br>Setting Audio Mode To: ${modeName}<br>Tap On [${buttonName}] Button To Select Other Modes.`);
    return;
  }

  applyMode();
}

function processCallEnd(callType) {
  if (config.debug) console.log(callType, 'Call Ended');

  // If you only want this to run in-call, stop monitoring when call ends.
  if (config.OnlyRunWhenInCall) stopMonitor();
}

/**************************************************************
 * Consolidated mic control
 **************************************************************/
function micKey({ ConnectorType, ConnectorId, SubId }) {
  return `${ConnectorType}.${ConnectorId}${SubId ? '.' + SubId : ''}`;
}

async function setGroupState(groupName, micList, micLevel /* 'duck' | 'unduck' */, force = false) {
  if (!force && groupState[groupName] === micLevel) return;
  const level = config.levels[micLevel];

  if (config.debug) {
    console.log(`Setting ${groupName} -> ${micLevel} (${gainLevel}=${level}) on: ${micList.map(micKey).join(', ')}`);
  }

  await Promise.all(micList.map(m => setInputLevelGain({ ...m, level })));
  groupState[groupName] = micLevel;
}

// Group: 'monitored' | 'controlled'
async function setMicsLevel(group, micLevel, force = false) {
  if (group === 'monitored') return setGroupState('monitoredMics', config.mics, micLevel, force);
  if (group === 'controlled') return setGroupState('controlledMics', config.duck, micLevel, force);
  throw new Error(`Invalid group "${group}". Use "monitored" or "controlled".`);
}

/**************************************************************
 * Audio event processing (includes upstream v1-3-0 timeout health logic)
 **************************************************************/
async function processAudioEvents(event) {
  // Clear any Audio Event Timeouts
  clearTimeout(audioEventTimeout);
  audioEventTimeout = null;

  // Count number of audio events for triggering periodic logging
  audioEventCount += 1;

  // Track number of missing events so we can restart VuMeter if required
  if (typeof event !== 'undefined' && Object.keys(event).length === 0) {
    missingAudioEventCounter += 1;
  } else {
    missingAudioEventCounter = 0;
  }

  const newLevels = flattenObject(event);

  // 1) Update per-mic sample buffers (micLevels)
  for (const [micName, levels] of Object.entries(micLevels)) {
    micLevels[micName].shift();
    micLevels[micName].push(
      (typeof newLevels?.[micName] === 'number')
        ? newLevels[micName]
        : levels[levels.length - 1]
    );
  }

  // 2) Compute rolling averages per mic AND store them into micLevelAverages history
  let aboveHighThreshold = false;
  let aboveLowThreshold = false;

  for (const [micName, levels] of Object.entries(micLevels)) {
    const sum = levels.reduce((partialSum, a) => partialSum + a, 0);
    const average = sum / levels.length;

    if (micLevelAverages?.[micName]) {
      micLevelAverages[micName].shift();
      micLevelAverages[micName].push(average);
    }

    if (!aboveHighThreshold && average > config.threshold.high) aboveHighThreshold = true;
    if (!aboveLowThreshold && average > config.threshold.low) aboveLowThreshold = true;
  }

  // 3) Ducking logic
  if (mode === 'autoDuck') {
    if (aboveHighThreshold) {
      clearTimeout(unduckTimer);
      unduckTimer = null;
      await setMicsLevel('controlled', 'duck');
    }

    if (!aboveLowThreshold && !unduckTimer) {
      unduckTimer = setTimeout(async () => {
        if (config.debug) console.log('Unducking Timeout Triggered');
        await setMicsLevel('controlled', 'unduck');
      }, config.unduck.timeout * 1000);
    }
  }

  // 4) Log real averages per mic every X audio events
  if (audioEventCount >= averageLogFequency) {
    if (config.debug) {
      console.log('Audio Level Stats (per mic)\n' + formatMicStatsForLog(micLevelAverages, micLevels));
    }
    audioEventCount = 0;
  }

  // 5) Monitor health
  if (missingAudioEventCounter >= maxMissingAudioEvents) {
    if (config.debug) console.log('Max Missed Audio Events Reached - Restarting Audio Monitor');
    startMonitor();
    return; // avoid double timeouts
  }

  startAudioEventTimeout();
}

/**************************************************************
 * Monitoring start/stop
 **************************************************************/
function startMonitor() {
  if (listener) {
    listener();
    listener = () => void 0;
  }

  listener = xapi.Event.Audio.Input.Connectors.on(processAudioEvents);

  const monitoringMicNames = startVuMeterConnectors.map(({ ConnectorType, ConnectorId }) => `${ConnectorType}.${ConnectorId}`);
  if (config.debug) console.log('Starting Audio Monitor:', ...monitoringMicNames);

  missingAudioEventCounter = 0;
  micLevelAverages = createMicLevels(averageLogFequency);
  micLevels = createMicLevels(config.samples);

  startVuMeterConnectors.forEach(({ ConnectorId, ConnectorType }) => {
    // Upstream compatibility: omit ConnectorId when not used/available
    if (typeof ConnectorId !== 'undefined' && ConnectorId !== null) {
      xapi.Command.Audio.VuMeter.Start({ ConnectorId, ConnectorType, IntervalMs: sampleInterval, Source: vuMeterSource });
    } else {
      xapi.Command.Audio.VuMeter.Start({ ConnectorType, IntervalMs: sampleInterval, Source: vuMeterSource });
    }
  });

  startAudioEventTimeout();
}

function stopMonitor() {
  if (config.debug) console.log('Stopping Audio Monitor');
  clearTimeout(audioEventTimeout);

  if (listener) {
    listener();
    listener = () => void 0;
  }

  xapi.Command.Audio.VuMeter.StopAll();
}

function startAudioEventTimeout() {
  clearTimeout(audioEventTimeout);
  const timeout = 2 * sampleInterval;

  if (config.debug) console.debug('Starting Audio Event Timeout - delay:', timeout, 'ms');

  audioEventTimeout = setTimeout(() => {
    if (config.debug) console.debug('Audio Event Timeout Reached - Triggering Audio Events Process');
    // Trigger empty event to maintain state machine and allow unducking
    processAudioEvents({});
  }, timeout);
}

/**************************************************************
 * Audio I/O helpers
 **************************************************************/
async function checkGainLevel() {
  const inputs = await xapi.Config.Audio.Input.get();
  const { Ethernet, Microphone, USBMicrophone } = inputs;

  if (Ethernet) return (typeof Ethernet?.[0]?.Channel?.[0]?.Gain !== 'undefined') ? 'Gain' : 'Level';
  if (Microphone) return (Microphone.some(mic => typeof mic?.Gain !== 'undefined')) ? 'Gain' : 'Level';
  if (USBMicrophone) return (typeof USBMicrophone?.[0]?.Gain !== 'undefined') ? 'Gain' : 'Level';
  return 'Gain';
}

async function setInputLevelGain({ ConnectorType, ConnectorId, SubId, level }) {
  const supportedTypes = ['Ethernet', 'Microphone', 'USBInterface', 'USBMicrophone'];
  if (!supportedTypes.includes(ConnectorType)) throw new Error(`Unsupported Audio Input Type [${ConnectorType}]`);

  try {
    if (SubId) {
      await xapi.Config.Audio.Input[ConnectorType][ConnectorId].Channel[SubId][gainLevel].set(level);
    } else {
      await xapi.Config.Audio.Input[ConnectorType][ConnectorId][gainLevel].set(level);
    }

    if (config.debug) {
      const mic = `${ConnectorType}.${ConnectorId}${SubId ? '.' + SubId : ''}`;
      console.log(`Mic: ${mic} - ${gainLevel}: ${level}`);
    }
  } catch (e) {
    // Swallow errors for devices/modes where some connectors are absent
    if (config.debug) console.warn('Failed to set input level/gain:', e?.message ?? e);
  }
}

function flattenObject(obj) {
  let result = {};

  for (const i in obj) {
    if (!Object.prototype.hasOwnProperty.call(obj, i)) continue;

    if (typeof obj[i] === 'object' && obj[i] !== null) {
      const flatObject = flattenObject(obj[i]);

      for (const x in flatObject) {
        if (!Object.prototype.hasOwnProperty.call(flatObject, x)) continue;

        const id = obj[i]?.id ?? i;
        const key = (x === 'VuMeter') ? id : ((i === 'SubId') ? x : id + '.' + x);
        result[key] = flatObject[x];
      }
    } else {
      if (i !== 'VuMeter') continue;
      result[i] = parseInt(obj[i], 10);
    }
  }

  return result;
}

/**************************************************************
 * UI helpers
 **************************************************************/
function alert(Text = '', Duration = 10) {
  if (!config.showAlerts) return;
  xapi.Command.UserInterface.Message.Alert.Display({ Duration, Target: 'Controller', Text, Title: config.button.name });
}

async function createPanel() {
  const { icon, color, name, location } = config.button;
  const panelId = config.panelId;
  const order = await panelOrder(panelId);

  const values = Object.keys(config.modeNames)
    .map(k => `<Value><Key>${k}</Key><Name>${escapeXml(config.modeNames[k])}</Name></Value>`)
    .join('');

  // If MTR is installed, panel is forced into ControlPanel unless Hidden
  const mtrDevice = await xapi.Command.MicrosoftTeams.List({ Show: 'Installed' })
    .then(() => true)
    .catch(() => false);

  const panelLocation = mtrDevice ? (location === 'Hidden' ? location : 'ControlPanel') : location;

  const panel = `
<Extensions>
  <Panel>
    <Origin>local</Origin>
    <Location>${panelLocation}</Location>
    <Icon>${escapeXml(icon)}</Icon>
    <Color>${escapeXml(color)}</Color>
    <Name>${escapeXml(name)}</Name>
    ${order}
    <ActivityType>Custom</ActivityType>
    <Page>
      <Name>${escapeXml(name)}</Name>
      <Row>
        <Widget>
          <WidgetId>${panelId}</WidgetId>
          <Type>GroupButton</Type>
          <Options>size=4;columns=1</Options>
          <ValueSpace>
            ${values}
          </ValueSpace>
        </Widget>
      </Row>
      <Options>hideRowNames=1</Options>
    </Page>
  </Panel>
</Extensions>`;

  return xapi.Command.UserInterface.Extensions.Panel.Save({ PanelId: panelId }, panel);
}

async function panelOrder(panelId) {
  const list = await xapi.Command.UserInterface.Extensions.List({ ActivityType: 'Custom' });
  const panels = list?.Extensions?.Panel;
  if (!panels) return '';

  const existingPanel = panels.find(p => p.PanelId === panelId);
  if (!existingPanel) return '';

  return `<Order>${existingPanel.Order}</Order>`;
}

function escapeXml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**************************************************************
 * List helpers
 **************************************************************/
function consolidateConnectors(inputArray) {
  const uniqueConnectors = new Map();

  inputArray.forEach(item => {
    const key = `${item.ConnectorType}-${item.ConnectorId}`;
    if (!uniqueConnectors.has(key)) {
      uniqueConnectors.set(key, { ConnectorType: item.ConnectorType, ConnectorId: item.ConnectorId });
    }
  });

  return Array.from(uniqueConnectors.values());
}

function createMicStrings(inputArray) {
  const uniqueConnectors = new Map();

  inputArray.forEach(({ ConnectorType, ConnectorId, SubId }) => {
    const key = `${ConnectorType}-${ConnectorId}-${SubId ?? ''}`;
    if (!uniqueConnectors.has(key)) uniqueConnectors.set(key, { ConnectorType, ConnectorId, SubId });
  });

  return Array.from(uniqueConnectors.values())
    .map(({ ConnectorType, ConnectorId, SubId }) => `${ConnectorType}.${ConnectorId}${SubId ? '.' + SubId : ''}`);
}

/**************************************************************
 * Format helpers
 **************************************************************/
function formatMicStatsForLog(averagesObj, rawLevelsObj) {
  const lines = [];

  for (const micName of Object.keys(averagesObj || {})) {
    const avgHistory = (averagesObj[micName] || []).filter(v => typeof v === 'number' && !Number.isNaN(v));
    const rawHistory = (rawLevelsObj[micName] || []).filter(v => typeof v === 'number' && !Number.isNaN(v));

    const avgCount = avgHistory.length;
    const latestAvg = avgCount ? avgHistory[avgCount - 1] : 0;
    const avgMin = avgCount ? Math.min(...avgHistory) : 0;
    const avgMax = avgCount ? Math.max(...avgHistory) : 0;

    const rawCount = rawHistory.length;
    const min = rawCount ? Math.min(...rawHistory) : 0;
    const max = rawCount ? Math.max(...rawHistory) : 0;

    lines.push(
      `${micName}: ` +
      `latestAvg=${latestAvg.toFixed(1)} ` +
      `AvgMin=${avgMin.toFixed(1)} ` +
      `AvgMax=${avgMax.toFixed(1)} ` +
      `Min=${min.toFixed(1)} ` +
      `Max=${max.toFixed(1)} ` +
      `(AverageLogFrequency=${avgCount}, Samples=${rawCount})`
    );
  }

  return lines.join('\n');
}
