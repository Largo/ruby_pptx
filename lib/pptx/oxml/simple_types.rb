# frozen_string_literal: true

require "pptx/errors"
require "pptx/length"

module Pptx
  module Oxml
    # Scalar types that appear as XML attribute values.
    #
    # Each is a class exposing `from_xml` / `to_xml`, mirroring the simple
    # types in the ECMA-376 schema. `to_xml` validates before converting, so an
    # out-of-range value raises here rather than producing a file PowerPoint
    # refuses to open.
    module SimpleTypes
      # Base for every simple type. Subclasses override `convert_from_xml`,
      # `convert_to_xml` and `validate`; the class-method inheritance mirrors
      # the Python original exactly.
      class BaseSimpleType
        class << self
          def from_xml(xml_value) = convert_from_xml(xml_value)

          def to_xml(value)
            value = normalize(value)
            validate(value)
            convert_to_xml(value)
          end

          # python-pptx's `Emu` is an `int` subclass, so a length could be
          # passed anywhere an integer was expected. Our {Pptx::Length} is a
          # separate class, so unwrap it here and every type gains the same
          # affordance: `shape.width = Pptx.inches(1)` works.
          def normalize(value) = value.is_a?(Pptx::Length) ? value.emu : value

          # Python's `round()` is round-half-to-even; Ruby's `Float#round`
          # defaults to half-up. Every rounding conversion here goes through
          # this helper so values landing exactly on .5 match python-pptx.
          def round_half_even(value) = value.round(half: :even)

          def validate_float(value)
            return if value.is_a?(Numeric) && !value.is_a?(Complex)

            raise TypeError, "value must be a number, got #{value.class}"
          end

          def validate_int(value)
            return if value.is_a?(Integer)

            raise TypeError, "value must be an integral type, got #{value.class}"
          end

          def validate_float_in_range(value, min_inclusive, max_inclusive)
            validate_float(value)
            return if value >= min_inclusive && value <= max_inclusive

            raise RangeError,
                  "value must be in range #{min_inclusive} to #{max_inclusive} inclusive, " \
                  "got #{value}"
          end

          def validate_int_in_range(value, min_inclusive, max_inclusive)
            validate_int(value)
            return if value >= min_inclusive && value <= max_inclusive

            raise RangeError,
                  "value must be in range #{min_inclusive} to #{max_inclusive} inclusive, " \
                  "got #{value}"
          end

          def validate_string(value)
            return value if value.is_a?(String)

            raise TypeError, "value must be a string, got #{value.class}"
          end
        end
      end

      class BaseFloatType < BaseSimpleType
        class << self
          def convert_from_xml(str_value) = Float(str_value)
          def convert_to_xml(value) = Float(value).to_s
          def validate(value) = validate_float(value)
        end
      end

      class BaseIntType < BaseSimpleType
        class << self
          def convert_from_percent_literal(str_value) = Integer(str_value.delete("%"), 10)
          def convert_from_xml(str_value) = Integer(str_value, 10)
          def convert_to_xml(value) = value.to_i.to_s
          def validate(value) = validate_int(value)
        end
      end

      class BaseStringType < BaseSimpleType
        class << self
          def convert_from_xml(str_value) = str_value
          def convert_to_xml(value) = value
          def validate(value) = validate_string(value)
        end
      end

      # A string type restricted to an enumerated set of values.
      class BaseStringEnumerationType < BaseStringType
        class << self
          # Declare the permitted values, each also defined as a constant.
          #
          #   values BAR: "bar", COL: "col"
          def values(**mapping)
            mapping.each { |name, value| const_set(name, value) }
            @members = mapping.values.freeze
          end

          def members = @members || superclass.members

          def validate(value)
            validate_string(value)
            return if members.include?(value)

            raise ArgumentError, "must be one of #{members.inspect}, got #{value.inspect}"
          end
        end
      end

      # -- xsd built-ins ---------------------------------------------------

      # Not validated against the URI grammar; the payoff does not justify it.
      class XsdAnyUri < BaseStringType; end

      class XsdBoolean < BaseSimpleType
        class << self
          def convert_from_xml(str_value)
            unless %w[1 0 true false].include?(str_value)
              raise InvalidXmlError,
                    "value must be one of '1', '0', 'true' or 'false', got #{str_value.inspect}"
            end

            %w[1 true].include?(str_value)
          end

          def convert_to_xml(value) = value ? "1" : "0"

          def validate(value)
            return if [true, false].include?(value)

            raise TypeError, "only true or false may be assigned, got #{value.inspect}"
          end
        end
      end

      class XsdDouble < BaseFloatType; end

      # Must start with a letter or underscore and contain no colon. Not fully
      # validated because it is not reachable from the public API.
      class XsdId < BaseStringType; end

      class XsdInt < BaseIntType
        def self.validate(value) = validate_int_in_range(value, -2_147_483_648, 2_147_483_647)
      end

      class XsdLong < BaseIntType
        def self.validate(value)
          validate_int_in_range(value, -9_223_372_036_854_775_808, 9_223_372_036_854_775_807)
        end
      end

      class XsdString < BaseStringType; end
      class XsdStringEnumeration < BaseStringEnumerationType; end

      # xsd:string with whitespace collapsed.
      class XsdToken < BaseStringType; end
      class XsdTokenEnumeration < BaseStringEnumerationType; end

      class XsdUnsignedByte < BaseIntType
        def self.validate(value) = validate_int_in_range(value, 0, 255)
      end

      class XsdUnsignedInt < BaseIntType
        def self.validate(value) = validate_int_in_range(value, 0, 4_294_967_295)
      end

      class XsdUnsignedShort < BaseIntType
        def self.validate(value) = validate_int_in_range(value, 0, 65_535)
      end

      # -- ST_* types from the ECMA-376 schema ------------------------------

      # `rot` on `<a:xfrm>`: 60000ths of a degree, exposed as float degrees.
      class ST_Angle < XsdInt
        DEGREE_INCREMENTS = 60_000
        THREE_SIXTY = 360 * DEGREE_INCREMENTS

        class << self
          def convert_from_xml(str_value)
            (Integer(str_value, 10) % THREE_SIXTY).to_f / DEGREE_INCREMENTS
          end

          # Normalized to a positive value; the modulo absorbs negative and
          # greater-than-360 inputs.
          def convert_to_xml(value)
            ((round_half_even(value * DEGREE_INCREMENTS)) % THREE_SIXTY).to_s
          end

          def validate(value) = BaseFloatType.validate(value)
        end
      end

      # `val` on c:majorUnit and friends.
      class ST_AxisUnit < XsdDouble
        def self.validate(value)
          super
          raise RangeError, "must be positive numeric value, got #{value}" if value <= 0.0
        end
      end

      class ST_BarDir < XsdStringEnumeration
        values BAR: "bar", COL: "col"
      end

      # Integer percent in 0..300, optionally written with a '%' suffix.
      class ST_BubbleScale < BaseIntType
        class << self
          def convert_from_xml(str_value)
            str_value.include?("%") ? convert_from_percent_literal(str_value) : super
          end

          def validate(value) = validate_int_in_range(value, 0, 300)
        end
      end

      # The schema gives a fierce regular expression; matching it would catch
      # programming errors only, so it is not enforced.
      class ST_ContentType < XsdString; end

      # xsd:union of ST_CoordinateUnqualified and ST_UniversalMeasure.
      class ST_Coordinate < BaseSimpleType
        class << self
          def convert_from_xml(str_value)
            if str_value.match?(/[imp]/)
              ST_UniversalMeasure.convert_from_xml(str_value)
            else
              Pptx::Length.emu(Integer(str_value, 10))
            end
          end

          def convert_to_xml(value) = value.to_i.to_s
          def validate(value) = ST_CoordinateUnqualified.validate(value)
        end
      end

      # xsd:union of ST_Coordinate32Unqualified and ST_UniversalMeasure.
      class ST_Coordinate32 < BaseSimpleType
        class << self
          def convert_from_xml(str_value)
            if str_value.match?(/[imp]/)
              ST_UniversalMeasure.convert_from_xml(str_value)
            else
              ST_Coordinate32Unqualified.convert_from_xml(str_value)
            end
          end

          def convert_to_xml(value) = ST_Coordinate32Unqualified.convert_to_xml(value)
          def validate(value) = ST_Coordinate32Unqualified.validate(value)
        end
      end

      class ST_Coordinate32Unqualified < XsdInt
        def self.convert_from_xml(str_value) = Pptx::Length.emu(Integer(str_value, 10))
      end

      class ST_CoordinateUnqualified < XsdLong
        def self.validate(value) = validate_int_in_range(value, -27_273_042_329_600, 27_273_042_316_900)
      end

      # `orient` on `<p:ph>`.
      class ST_Direction < XsdTokenEnumeration
        values HORZ: "horz", VERT: "vert"
      end

      class ST_DrawingElementId < XsdUnsignedInt; end

      class ST_Extension < XsdString; end

      # Integer percent in 0..500, optionally written with a '%' suffix.
      class ST_GapAmount < BaseIntType
        class << self
          def convert_from_xml(str_value)
            str_value.include?("%") ? convert_from_percent_literal(str_value) : super
          end

          def validate(value) = validate_int_in_range(value, 0, 500)
        end
      end

      # `val` on <c:grouping>; doubles as ST_BarGrouping, which shares the tag.
      class ST_Grouping < XsdStringEnumeration
        values CLUSTERED: "clustered", PERCENT_STACKED: "percentStacked",
               STACKED: "stacked", STANDARD: "standard"
      end

      class ST_HexColorRGB < BaseStringType
        class << self
          # Upper-cased purely for consistency of output.
          def convert_to_xml(value) = value.upcase

          def validate(value)
            str_value = validate_string(value)
            unless str_value.length == 6
              raise ArgumentError, "RGB string must be six characters long, got #{str_value.inspect}"
            end
            unless str_value.match?(/\A\h{6}\z/)
              raise ArgumentError, "RGB string must be a valid hex string, got #{str_value.inspect}"
            end
          end
        end
      end

      # `val` on c:xMode and other CT_LayoutMode elements.
      class ST_LayoutMode < XsdStringEnumeration
        values EDGE: "edge", FACTOR: "factor"
      end

      # 0..1000 inclusive, optionally written with a '%' suffix.
      class ST_LblOffset < XsdUnsignedShort
        class << self
          def convert_from_xml(str_value)
            str_value.end_with?("%") ? convert_from_percent_literal(str_value) : Integer(str_value, 10)
          end

          def validate(value) = validate_int_in_range(value, 0, 1000)
        end
      end

      class ST_LineWidth < XsdInt
        class << self
          def convert_from_xml(str_value) = Pptx::Length.emu(Integer(str_value, 10))

          def validate(value)
            super
            return if value >= 0 && value <= 20_116_800

            raise RangeError,
                  "value must be in range 0-20116800 inclusive (0-1584 points), got #{value}"
          end
        end
      end

      class ST_MarkerSize < XsdUnsignedByte
        def self.validate(value) = validate_int_in_range(value, 2, 72)
      end

      # `val` on c:orientation (CT_Orientation).
      class ST_Orientation < XsdStringEnumeration
        values MAX_MIN: "maxMin", MIN_MAX: "minMax"
      end

      # Integer percent in -100..100, optionally written with a '%' suffix.
      class ST_Overlap < BaseIntType
        class << self
          def convert_from_xml(str_value)
            str_value.include?("%") ? convert_from_percent_literal(str_value) : super
          end

          def validate(value) = validate_int_in_range(value, -100, 100)
        end
      end

      # Either an integer literal in 1000ths of a percent ("42000"), or a float
      # with a '%' suffix ("42.0%"). Exposed as a float fraction, so 0.42.
      class ST_Percentage < BaseIntType
        class << self
          def convert_from_xml(str_value)
            return convert_from_percent_literal_float(str_value) if str_value.include?("%")

            Integer(str_value, 10) / 100_000.0
          end

          def convert_to_xml(value) = round_half_even(value * 100_000.0).to_s

          def validate(value) = validate_float_in_range(value, -21_474.83648, 21_474.83647)

          def convert_from_percent_literal_float(str_value)
            Float(str_value.delete_suffix("%")) / 100.0
          end
        end
      end

      # `sz` on <p:ph>.
      class ST_PlaceholderSize < XsdTokenEnumeration
        values FULL: "full", HALF: "half", QUARTER: "quarter"
      end

      class ST_PositiveCoordinate < XsdLong
        class << self
          def convert_from_xml(str_value) = Pptx::Length.emu(super)
          def validate(value) = validate_int_in_range(value, 0, 27_273_042_316_900)
        end
      end

      # `a:lin@ang`: positive angles less than 360 degrees.
      class ST_PositiveFixedAngle < ST_Angle
        def self.convert_to_xml(degrees)
          if degrees < 0.0
            degrees = (degrees % -360) + 360
          elsif degrees > 0.0
            degrees %= 360
          end
          round_half_even(degrees * DEGREE_INCREMENTS).to_s
        end
      end

      # ST_Percentage constrained to 0.0..1.0.
      class ST_PositiveFixedPercentage < ST_Percentage
        def self.validate(value) = validate_float_in_range(value, 0.0, 1.0)
      end

      class ST_RelationshipId < XsdString; end

      class ST_SlideId < XsdUnsignedInt
        def self.validate(value) = validate_int_in_range(value, 256, 2_147_483_647)
      end

      # The 36 layout kinds ST_SlideLayoutType allows. PowerPoint uses this to
      # decide which layout to offer for a given command, so an unrecognised
      # value is rejected rather than written through.
      class ST_SlideLayoutType < XsdString
        VALUES = %w[
        title tx twoColTx tbl txAndChart chartAndTx dgm chart txAndClipArt clipArtAndTx
        titleOnly blank txAndObj objAndTx objOnly obj txAndMedia mediaAndTx objOverTx txOverObj
        txAndTwoObj twoObjAndTx twoObjOverTx fourObj vertTx clipArtAndVertTx vertTitleAndTx
        vertTitleAndTxOverChart twoObj objAndTwoObj twoObjAndObj cust secHead twoTxTwoObj objTx
        picTx
        ].freeze

        def self.validate(value)
          super
          return if VALUES.include?(value)

          raise ArgumentError, "#{value.inspect} is not a slide layout type"
        end
      end

      # Slide-master and slide-layout ids share one range, which starts above
      # the signed 32-bit maximum -- PowerPoint numbers them from 2147483648
      # upwards precisely so they cannot be confused with slide ids.
      class ST_SlideMasterId < XsdUnsignedInt
        def self.validate(value) = validate_int_in_range(value, 2_147_483_648, 4_294_967_295)
      end

      class ST_SlideLayoutId < ST_SlideMasterId; end

      class ST_SlideSizeCoordinate < BaseIntType
        class << self
          def convert_from_xml(str_value) = Pptx::Length.emu(Integer(str_value, 10))

          def validate(value)
            validate_int(value)
            return if value >= 914_400 && value <= 51_206_400

            raise RangeError,
                  "value must be in range 914400 to 51206400 (1-56 inches), got #{value}"
          end
        end
      end

      class ST_Style < XsdUnsignedByte
        def self.validate(value) = validate_int_in_range(value, 1, 48)
      end

      # `TargetMode` on a Relationship element.
      class ST_TargetMode < XsdString
        EXTERNAL = "External"
        INTERNAL = "Internal"

        def self.validate(value)
          validate_string(value)
          return if [EXTERNAL, INTERNAL].include?(value)

          raise ArgumentError, "must be one of 'Internal' or 'External', got #{value.inspect}"
        end
      end

      # `fontScale` on <a:normAutofit>, as a float percent.
      class ST_TextFontScalePercentOrPercentString < BaseFloatType
        class << self
          def convert_from_xml(str_value)
            return Float(str_value.delete_suffix("%")) if str_value.end_with?("%")

            Integer(str_value, 10) / 1000.0
          end

          def convert_to_xml(value) = (value * 1000.0).to_i.to_s

          def validate(value)
            BaseFloatType.validate(value)
            return if value >= 1.0 && value <= 100.0

            raise RangeError, "value must be in range 1.0..100.0 (percent), got #{value}"
          end
        end
      end

      class ST_TextFontSize < BaseIntType
        def self.validate(value) = validate_int_in_range(value, 100, 400_000)
      end

      class ST_TextIndentLevelType < BaseIntType
        def self.validate(value) = validate_int_in_range(value, 0, 8)
      end

      # Line spacing as a multiple of line height, so 1.75 <-> "175000".
      class ST_TextSpacingPercentOrPercentString < BaseFloatType
        class << self
          def convert_from_xml(str_value)
            return Float(str_value.delete_suffix("%")) / 100.0 if str_value.end_with?("%")

            Integer(str_value, 10) / 100_000.0
          end

          def convert_to_xml(value) = round_half_even(value * 100_000.0).to_s

          def validate(value) = validate_float_in_range(value, 0.0, 132.0)
        end
      end

      # Reads centipoints, exposes a {Pptx::Length}.
      class ST_TextSpacingPoint < BaseIntType
        class << self
          def convert_from_xml(str_value) = Pptx::Length.centipoints(Integer(str_value, 10))
          def convert_to_xml(value) = Pptx::Length.emu(value).centipoints.to_s
          def validate(value) = validate_int_in_range(value, 0, 20_116_800)
        end
      end

      class ST_TextTypeface < XsdString; end

      # `wrap` on <a:bodyPr>.
      class ST_TextWrappingType < XsdTokenEnumeration
        values NONE: "none", SQUARE: "square"
      end

      # A float with a two-character unit suffix, e.g. "1.5in".
      class ST_UniversalMeasure < BaseSimpleType
        MULTIPLIERS = {
          "mm" => 36_000, "cm" => 360_000, "in" => 914_400,
          "pt" => 12_700, "pc" => 152_400, "pi" => 152_400
        }.freeze

        class << self
          def convert_from_xml(str_value)
            units = str_value[-2..]
            multiplier = MULTIPLIERS.fetch(units) do
              raise InvalidXmlError, "unknown measurement unit #{units.inspect} in #{str_value.inspect}"
            end
            Pptx::Length.emu(round_half_even(Float(str_value[0...-2]) * multiplier))
          end

          def convert_to_xml(value) = value.to_i.to_s
          def validate(value) = validate_int(value)
        end
      end
    end
  end
end
