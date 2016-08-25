/*(c) Copyright 2015 Pivotal Software, Inc. All Rights Reserved.*/
'use strict';

var _extends = Object.assign || function (target) { for (var i = 1; i < arguments.length; i++) { var source = arguments[i]; for (var key in source) { if (Object.prototype.hasOwnProperty.call(source, key)) { target[key] = source[key]; } } } return target; };

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

var _mixins = require('../mixins');

var _mixins2 = _interopRequireDefault(_mixins);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var raf = require('raf');
var React = require('react');
var ReactDOM = require('react-dom');
var shallowEqual = require('fbjs/lib/shallowEqual');

var ShallowCompare = require('../mixins/shallow_compare_mixin');
var Component = (0, _mixins2.default)(React.Component).with(ShallowCompare);

var rafify = function rafify(callback) {
  return function () {
    var _this = this;

    for (var _len = arguments.length, args = Array(_len), _key = 0; _key < _len; _key++) {
      args[_key] = arguments[_key];
    }

    raf(function () {
      return callback.call.apply(callback, [_this].concat(args));
    });
  };
};

var privates = new WeakMap();

var properties = ['width', 'height', 'top', 'right', 'bottom', 'left'];

module.exports = {
  useBoundingClientRect: function useBoundingClientRect(Klass) {
    return function (_Component) {
      _inherits(BoundingClientRect, _Component);

      function BoundingClientRect(props, context) {
        _classCallCheck(this, BoundingClientRect);

        var _this2 = _possibleConstructorReturn(this, Object.getPrototypeOf(BoundingClientRect).call(this, props, context));

        _this2.resize = function () {
          var _ref = privates.get(_this2) || {};

          var prevBoundingClientRect = _ref.boundingClientRect;

          var boundingClientRect = _this2.getBoundingClientRect();
          var isNotEqual = function isNotEqual(property) {
            return boundingClientRect[property] !== prevBoundingClientRect[property];
          };
          if (!prevBoundingClientRect || properties.some(isNotEqual)) _this2.forceUpdate();
        };

        var resolver = void 0;
        var containerReady = new Promise(function (resolve) {
          return resolver = resolve;
        });
        containerReady.resolve = resolver;
        var state = _this2.state;

        _this2.state = _extends({}, state, { container: null, containerReady: containerReady });
        _this2.resize = rafify(_this2.resize);

        _this2.getBoundingClientRect = _this2.getBoundingClientRect.bind(_this2);
        return _this2;
      }

      _createClass(BoundingClientRect, [{
        key: 'componentDidMount',
        value: function componentDidMount() {
          var _this3 = this;

          privates.set(this, { resize: this.resize });
          window.addEventListener('resize', this.resize);
          this.setState({ container: ReactDOM.findDOMNode(this.component) });
          setImmediate(function () {
            return _this3.state.containerReady.resolve(_this3.state.container);
          });
        }
      }, {
        key: 'componentWillUnmount',
        value: function componentWillUnmount() {
          var _ref2 = privates.get(this) || {};

          var resize = _ref2.resize;

          window.removeEventListener('resize', resize);
          privates.delete(this);
        }
      }, {
        key: 'componentWillReceiveProps',
        value: function componentWillReceiveProps(nextProps) {
          if (!shallowEqual(this.props, nextProps)) this.resize();
        }
      }, {
        key: 'getBoundingClientRect',
        value: function getBoundingClientRect() {
          return this.state.container && this.state.container.getBoundingClientRect() || {};
        }
      }, {
        key: 'render',
        value: function render() {
          var _this4 = this;

          var _ref3 = privates.get(this) || {};

          var resize = _ref3.resize;

          var boundingClientRect = this.getBoundingClientRect();
          privates.set(this, { boundingClientRect: boundingClientRect, resize: resize });
          return React.createElement(Klass, _extends({}, this.props, this.state, { boundingClientRect: boundingClientRect }, { ref: function ref(_ref4) {
              return _this4.component = _ref4;
            } }));
        }
      }]);

      return BoundingClientRect;
    }(Component);
  }
};