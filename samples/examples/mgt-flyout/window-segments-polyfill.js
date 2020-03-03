//https://github.com/zouhir/windowsegments-polyfill
const POLYFILL_NAMESPACE = '__foldables_env_vars__';
const SPANNING_MF_VAL_HOR = 'single-fold-horizontal';
const SPANNING_MF_VAL_VER = 'single-fold-vertical';
const SPANNING_MF_VAL_NONE = 'none';

if (typeof window[POLYFILL_NAMESPACE] === 'undefined') {
    // polyfill configuration related variables
    const spanning = sessionStorage.getItem(`${POLYFILL_NAMESPACE}-spanning`) || SPANNING_MF_VAL_NONE;
    const foldSize = sessionStorage.getItem(`${POLYFILL_NAMESPACE}-foldSize`) || 0;
    const browserShellSize = sessionStorage.getItem(`${POLYFILL_NAMESPACE}-browserShellSize`) || 0;
    // global configs, accessible via the window object
    Object.defineProperty(window, POLYFILL_NAMESPACE, {
        value: {
            spanning,
            foldSize,
            browserShellSize,
            update,
            onupdate: [onDeviceSettingsUpdated],
        },
    });

    // web-based emulator runs this polyfill in an iframe, we need to communicate
    // emulator state changes to the site
    // should only be registered once (in CSS or JS polyfill not both)
    window.addEventListener('message', evt => {
        const action = evt.data.action || '';
        const value = evt.data.value || {};
        if (action === 'update') {
            window[POLYFILL_NAMESPACE].update(value);
        }
    });
} else {
    window[POLYFILL_NAMESPACE].onupdate.push(onDeviceSettingsUpdated);
}

function getWindowSegments() {
    const segments = [];
    const viewportWidth = window.innerWidth;
    const viewporHeight = window.innerHeight;

    let { foldSize, browserShellSize, spanning } = window[POLYFILL_NAMESPACE];

    // those are numbers not strings as stored in sessionStorage
    foldSize = +foldSize;
    browserShellSize = +browserShellSize;

    if (spanning === SPANNING_MF_VAL_NONE) {
        segments.push({
            left: 0,
            top: 0,
            width: viewportWidth,
            height: viewporHeight,
        });
    }

    if (spanning === SPANNING_MF_VAL_VER) {
        const foldLeft = viewportWidth / 2 - foldSize / 2;
        segments.push(
            {
                top: 0,
                left: 0,
                width: foldLeft,
                height: viewporHeight,
            },
            {
                left: foldLeft + foldSize,
                top: 0,
                width: viewportWidth - (foldLeft + foldSize),
                height: viewporHeight,
            },
        );
    }

    if (spanning === SPANNING_MF_VAL_HOR) {
        const foldTop = window.innerHeight / 2 - foldSize / 2 - browserShellSize;
        segments.push(
            {
                top: 0,
                left: 0,
                width: viewportWidth,
                height: foldTop,
            },
            {
                left: 0,
                top: foldTop + foldSize,
                width: viewportWidth,
                height: viewporHeight - foldTop + foldSize,
            },
        );
    }
    return segments;
}

if (typeof window.getWindowSegments === 'undefined') {
    Object.defineProperty(window, 'getWindowSegments', {
        value: getWindowSegments,
    });
}

function onDeviceSettingsUpdated() {
    console.warn('device settings updated, resize the window');
}

const VALID_CONFIG_PROPS = new Set(['foldSize', 'browserShellSize', 'spanning']);

function update(newConfings) {
    Object.keys(newConfings).forEach(k => {
        if (VALID_CONFIG_PROPS.has(k)) {
            window[POLYFILL_NAMESPACE][k] = newConfings[k];
            sessionStorage.setItem(`${POLYFILL_NAMESPACE}-${k}`, window[POLYFILL_NAMESPACE][k]);
        }
    });
    window[POLYFILL_NAMESPACE].onupdate.forEach(fn => fn());
}
