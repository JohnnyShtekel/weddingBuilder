//(c) Copyright 2015 Pivotal Software, Inc. All Rights Reserved.
'use strict';

var _map = require('babel-runtime/core-js/map');

var _map2 = _interopRequireDefault(_map);

var _extends2 = require('babel-runtime/helpers/extends');

var _extends3 = _interopRequireDefault(_extends2);

var _slicedToArray2 = require('babel-runtime/helpers/slicedToArray');

var _slicedToArray3 = _interopRequireDefault(_slicedToArray2);

var _getIterator2 = require('babel-runtime/core-js/get-iterator');

var _getIterator3 = _interopRequireDefault(_getIterator2);

var _isNan = require('babel-runtime/core-js/number/is-nan');

var _isNan2 = _interopRequireDefault(_isNan);

var _weakMap = require('babel-runtime/core-js/weak-map');

var _weakMap2 = _interopRequireDefault(_weakMap);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var Easing = require('easing-js');
var now = require('performance-now');
var raf = require('raf');

var privates = new _weakMap2.default();

function isNumber(obj) {
  return typeof obj === 'number' && !(0, _isNan2.default)(obj);
}

function strip(number) {
  return parseFloat(number.toPrecision(12));
}

function someAnimating(animations) {
  var _iteratorNormalCompletion = true;
  var _didIteratorError = false;
  var _iteratorError = undefined;

  try {
    for (var _iterator = (0, _getIterator3.default)(animations), _step; !(_iteratorNormalCompletion = (_step = _iterator.next()).done); _iteratorNormalCompletion = true) {
      var _step$value = (0, _slicedToArray3.default)(_step.value, 2);

      var animation = _step$value[1];

      if (animation.isAnimating) return true;
    }
  } catch (err) {
    _didIteratorError = true;
    _iteratorError = err;
  } finally {
    try {
      if (!_iteratorNormalCompletion && _iterator.return) {
        _iterator.return();
      }
    } finally {
      if (_didIteratorError) {
        throw _iteratorError;
      }
    }
  }

  return false;
}

function scheduleAnimation(context) {
  raf(function () {
    var animations = privates.get(context);
    var currentTime = now();
    var shouldUpdate = false;
    animations && animations.forEach(function (animation, name) {
      var isFunction = typeof name === 'function';
      if (!animation.isAnimating) return;

      var duration = animation.duration;
      var easing = animation.easing;
      var endValue = animation.endValue;
      var startTime = animation.startTime;
      var startValue = animation.startValue;

      var deltaTime = currentTime - startTime;
      if (deltaTime >= duration) {
        (0, _extends3.default)(animation, { isAnimating: false, startTime: currentTime, value: endValue });
      } else {
        animation.value = strip(Easing[easing](deltaTime, startValue, endValue - startValue, duration));
      }

      shouldUpdate = shouldUpdate || !isFunction;
      if (isFunction) name(animation.value);
    });

    if (animations && someAnimating(animations)) scheduleAnimation(context);
    if (shouldUpdate) context.forceUpdate();
  });
}

var AnimationMixin = {
  componentWillUnmount: function componentWillUnmount() {
    privates.delete(this);
  },
  shouldAnimate: function shouldAnimate() {
    return true;
  },
  animate: function animate(name, endValue, duration) {
    var options = arguments.length <= 3 || arguments[3] === undefined ? {} : arguments[3];

    var animations = privates.get(this);
    if (!animations) {
      privates.set(this, animations = new _map2.default());
    }

    var animation = animations.get(name);
    var shouldAnimate = this.shouldAnimate() && options.animation !== false;
    if (!animation || !shouldAnimate || !isNumber(endValue)) {
      var easing = options.easing || 'linear';
      var startValue = isNumber(options.startValue) && shouldAnimate ? options.startValue : endValue;
      animation = { duration: duration, easing: easing, endValue: endValue, isAnimating: false, startValue: startValue, value: startValue };
      animations.set(name, animation);
    }

    if (!duration) {
      (0, _extends3.default)(animation, { endValue: endValue, value: endValue });
      animations.set(name, animation);
    }

    if (animation.value !== endValue && !animation.isAnimating) {
      if (!someAnimating(animations)) scheduleAnimation(this);
      var startTime = 'startTime' in options ? options.startTime : now();
      duration = duration || animation.duration;
      var easing = options.easing || animation.easing;
      var startValue = animation.value;
      (0, _extends3.default)(animation, { isAnimating: true, endValue: endValue, startValue: startValue, startTime: startTime, duration: duration, easing: easing });
    }

    return animation.value;
  }
};

module.exports = AnimationMixin;