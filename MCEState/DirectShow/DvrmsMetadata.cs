// Stephen Toub
// stoub@microsoft.com

using System;
using System.IO;
using System.Text;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.DirectShow;

namespace Microsoft.DirectShow.Metadata
{
    /// <summary>
    /// Metadata editor for DVR-MS files.
    /// </summary>
    public sealed class DvrmsMetadata: IDisposable
	{
		IStreamBufferRecordingAttribute _editor;

		/// <summary>Initializes the editor.</summary>
		/// <param name="filepath">The path to the file.</param>
		public DvrmsMetadata(string filepath)
		{
			IFileSourceFilter sourceFilter = (IFileSourceFilter)ClassId.CoCreateInstance(ClassId.RecordingAttributes);
			sourceFilter.Load(filepath, null);
			_editor = (IStreamBufferRecordingAttribute)sourceFilter;
		}

		/// <summary>Gets all of the attributes on a file.</summary>
		/// <returns>A collection of the attributes from the file.</returns>
        public Dictionary<string, MetadataItem> GetAttributes()
		{
			if (_editor == null) throw new ObjectDisposedException(GetType().Name);

            Dictionary<string, MetadataItem> propsRetrieved = new Dictionary<string, MetadataItem>();
			object obj = _editor.EnumAttributes();

			// Get the number of attributes
			ushort attributeCount = _editor.GetAttributeCount(0);

			// Get each attribute by index
			for(ushort i = 0; i < attributeCount; i++)
			{
				MetadataItemType attributeType;
				StringBuilder attributeName = null;
				byte[] attributeValue = null;
				ushort attributeNameLength = 0;
				ushort attributeValueLength = 0;

				// Get the lengths of the name and the value, then use them to create buffers to receive them
				uint reserved = 0;
				_editor.GetAttributeByIndex(i, ref reserved, attributeName, ref attributeNameLength,
					out attributeType, attributeValue, ref attributeValueLength);
				attributeName = new StringBuilder(attributeNameLength);
				attributeValue = new byte[attributeValueLength];

				// Get the name and value
				_editor.GetAttributeByIndex(i, ref reserved, attributeName, ref attributeNameLength,
					out attributeType, attributeValue, ref attributeValueLength);

				// If we got a name, parse the value and add the metadata item
				if (attributeName != null && attributeName.Length > 0)
				{
					object val = ParseAttributeValue(attributeType, attributeValue);
					string key = attributeName.ToString().TrimEnd('\0');
					propsRetrieved[key] = new MetadataItem(key, val, attributeType);
				}
			}

			// Return the parsed items
			return propsRetrieved;
		}

        /// <summary>Gets the value of the specified attribute.</summary>
        /// <param name="itemType">The type of the attribute.</param>
        /// <param name="valueData">The byte array to be parsed.</param>
        private static object ParseAttributeValue(MetadataItemType itemType, byte[] valueData)
        {
            if (!Enum.IsDefined(typeof(MetadataItemType), itemType))
                throw new ArgumentOutOfRangeException("itemType");
            if (valueData == null) throw new ArgumentNullException("valueData");

            // Convert the attribute value to a byte array based on the item type.
            switch (itemType)
            {
                case MetadataItemType.String:
                    StringBuilder sb = new StringBuilder(valueData.Length);
                    for (int i = 0; i < valueData.Length - 2; i += 2)
                    {
                        sb.Append(Convert.ToString(BitConverter.ToChar(valueData, i)));
                    }
                    string result = sb.ToString();
                    if (result.EndsWith("\\0")) result = result.Substring(0, result.Length - 2);
                    return result;
                case MetadataItemType.Boolean: return BitConverter.ToBoolean(valueData, 0);
                case MetadataItemType.Dword: return BitConverter.ToInt32(valueData, 0);
                case MetadataItemType.Qword: return BitConverter.ToInt64(valueData, 0);
                case MetadataItemType.Word: return BitConverter.ToInt16(valueData, 0);
                case MetadataItemType.Guid: return new Guid(valueData);
                case MetadataItemType.Binary: return valueData;
                default: throw new ArgumentOutOfRangeException("itemType");
            }
        }

        /// <summary>
        /// Releases unmanaged and - optionally - managed resources
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Releases unmanaged and - optionally - managed resources
        /// </summary>
        /// <param name="disposeManagedObjs"><c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged resources.</param>
        void Dispose(bool disposeManagedObjs)
		{
            if (disposeManagedObjs && _editor != null)
			{
				while(Marshal.ReleaseComObject(_editor) > 0);
				_editor = null;
			}
		}

        /// <summary>
        /// Releases unmanaged resources and performs other cleanup operations before the
        /// <see cref="DvrmsMetadata"/> is reclaimed by garbage collection.
        /// </summary>
        ~DvrmsMetadata()
        {
            Dispose(false);
        }

		[ComImport]
		[Guid("16CA4E03-FE69-4705-BD41-5B7DFC0C95F3")]
		[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
		private interface IStreamBufferRecordingAttribute
		{
			/// <summary>Sets an attribute on a recording object. If an attribute of the same name already exists, overwrites the old.</summary>
			/// <param name="ulReserved">Reserved. Set this parameter to zero.</param>
			/// <param name="pszAttributeName">Wide-character string that contains the name of the attribute.</param>
			/// <param name="StreamBufferAttributeType">Defines the data type of the attribute data.</param>
			/// <param name="pbAttribute">Pointer to a buffer that contains the attribute data.</param>
			/// <param name="cbAttributeLength">The size of the buffer specified in pbAttribute.</param>
			void SetAttribute(
				[In] uint ulReserved, 
				[In, MarshalAs(UnmanagedType.LPWStr)] string pszAttributeName,
				[In] MetadataItemType StreamBufferAttributeType,
				[In, MarshalAs(UnmanagedType.LPArray)] byte [] pbAttribute,
				[In] ushort cbAttributeLength);

			/// <summary>Returns the number of attributes that are currently defined for this stream buffer file.</summary>
			/// <param name="ulReserved">Reserved. Set this parameter to zero.</param>
			/// <returns>Number of attributes that are currently defined for this stream buffer file.</returns>
			ushort GetAttributeCount([In] uint ulReserved);

			/// <summary>Given a name, returns the attribute data.</summary>
			/// <param name="pszAttributeName">Wide-character string that contains the name of the attribute.</param>
			/// <param name="pulReserved">Reserved. Set this parameter to zero.</param>
			/// <param name="pStreamBufferAttributeType">
			/// Pointer to a variable that receives a member of the STREAMBUFFER_ATTR_DATATYPE enumeration. 
			/// This value indicates the data type that you should use to interpret the attribute, which is 
			/// returned in the pbAttribute parameter.
			/// </param>
			/// <param name="pbAttribute">
			/// Pointer to a buffer that receives the attribute, as an array of bytes. Specify the size of the buffer in the 
			/// pcbLength parameter. To find out the required size for the array, set pbAttribute to NULL and check the 
			/// value that is returned in pcbLength.
			/// </param>
			/// <param name="pcbLength">
			/// On input, specifies the size of the buffer given in pbAttribute, in bytes. On output, 
			/// contains the number of bytes that were copied to the buffer.
			/// </param>
			void GetAttributeByName(
				[In, MarshalAs(UnmanagedType.LPWStr)] string pszAttributeName,
				[In] ref uint pulReserved,
				[Out] out MetadataItemType pStreamBufferAttributeType,
				[Out, MarshalAs(UnmanagedType.LPArray)] byte[] pbAttribute,
				[In, Out] ref ushort pcbLength);

			/// <summary>The GetAttributeByIndex method retrieves an attribute, specified by index number.</summary>
			/// <param name="wIndex">Zero-based index of the attribute to retrieve.</param>
			/// <param name="pulReserved">Reserved. Set this parameter to zero.</param>
			/// <param name="pszAttributeName">Pointer to a buffer that receives the name of the attribute, as a null-terminated wide-character string.</param>
			/// <param name="pcchNameLength">On input, specifies the size of the buffer given in pszAttributeName, in wide characters.</param>
			/// <param name="pStreamBufferAttributeType">Pointer to a variable that receives a member of the STREAMBUFFER_ATTR_DATATYPE enumeration.</param>
			/// <param name="pbAttribute">Pointer to a buffer that receives the attribute, as an array of bytes.</param>
			/// <param name="pcbLength">On input, specifies the size of the buffer given in pbAttribute, in bytes.</param>
			void GetAttributeByIndex (
				[In] ushort wIndex,
				[In, Out] ref uint pulReserved,
				[Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszAttributeName,
				[In, Out] ref ushort pcchNameLength,
				[Out] out MetadataItemType pStreamBufferAttributeType,
				[Out, MarshalAs(UnmanagedType.LPArray)] byte [] pbAttribute,
				[In, Out] ref ushort pcbLength);

			/// <summary>The EnumAttributes method enumerates the existing attributes of the stream buffer file.</summary>
			/// <returns>Address of a variable that receives an IEnumStreamBufferRecordingAttrib interface pointer.</returns>
			[return: MarshalAs(UnmanagedType.Interface)]
			object EnumAttributes();
		}

    }

    /// <summary>The type of a metadata attribute value.</summary>
    public enum MetadataItemType
    {
        /// <summary>DWORD</summary>
        Dword = 0,
        /// <summary>String</summary>
        String = 1,
        /// <summary>Binary</summary>
        Binary = 2,
        /// <summary>Boolean</summary>
        Boolean = 3,
        /// <summary>QWORD</summary>
        Qword = 4,
        /// <summary>WORD</summary>
        Word = 5,
        /// <summary>Guid</summary>
        Guid = 6,
    }

    /// <summary>Represents a metadata attribute.</summary>
    public class MetadataItem : ICloneable
    {
        /// <summary>The name of the attribute.</summary>
        private string _name;
        /// <summary>The value of the attribute.</summary>
        private object _value;
        /// <summary>The type of the attribute value.</summary>
        private MetadataItemType _type;

        /// <summary>Initializes the metadata item.</summary>
        /// <param name="name">The name of the attribute.</param>
        /// <param name="value">The value of the attribute.</param>
        /// <param name="type">The type of the attribute value.</param>
        internal MetadataItem(string name, object value, MetadataItemType type)
        {
            Name = name;
            Value = value;
            Type = type;
        }

        /// <summary>Gets or sets the name of the attribute.</summary>
        public string Name { get { return _name; } set { _name = value; } }
        /// <summary>Gets or sets the value of the attribute.</summary>
        public object Value { get { return _value; } set { _value = value; } }
        /// <summary>Gets or sets the type of the attribute value.</summary>
        public MetadataItemType Type { get { return _type; } set { _type = value; } }

        /// <summary>Clones the attribute item.</summary>
        /// <returns>A shallow copy of the attribute.</returns>
        public MetadataItem Clone() { return (MetadataItem)MemberwiseClone(); }

        /// <summary>Clones the attribute item.</summary>
        /// <returns>A shallow copy of the attribute.</returns>
        object ICloneable.Clone() { return Clone(); }
    }

}