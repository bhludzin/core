/*
 * (c) Copyright Ascensio System SIA 2010-2023
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */
#pragma once

#include "../../WrapperWritingElement.h"
#include "../../Limit/BlipCompression.h"
#include "../UniEffect.h"
#include "../ExtP.h"
#include "../../../DocxFormat/RId.h"

namespace PPTX
{
	namespace Logic
	{
		class Blip : public WrapperWritingElement
		{
		public:
			WritingElement_AdditionMethods(Blip)

			Blip(std::wstring ns = L"a");
			Blip& operator=(const Blip& oSrc);

			virtual OOX::EElementType getType() const;

			virtual void fromXML(XmlUtils::CXmlNode& node);
			virtual void fromXML(XmlUtils::CXmlLiteReader& oReader);
			
			void ReadAttributes(XmlUtils::CXmlLiteReader& oReader);

			virtual std::wstring toXML() const;
			virtual void toXmlWriter(NSBinPptxRW::CXmlWriter* pWriter) const;

			virtual void toPPTY(NSBinPptxRW::CBinaryFileWriter* pWriter) const;
			virtual void fromPPTY(NSBinPptxRW::CBinaryFileReader* pReader);
			
			virtual std::wstring GetFullPicName(OOX::IFileContainer* pRels = NULL) const;
			virtual std::wstring GetFullOleName(const OOX::RId& pRId, OOX::IFileContainer* pRels = NULL) const;

			std::vector<UniEffect>	Effects;
			std::vector<Ext>		ExtLst;

			// The a:ext entries of a:extLst that this model does not read (PowerPoint's a14:imgProps
			// picture corrections, whose a14:imgLayer relates the original picture as an HD Photo part,
			// a14:useLocalDpi, ...). They pass through the editor as the blip's PPTY record 4 and are
			// written back as read, with their parts, so PowerPoint can still change or reset them.
			struct UnknownExt
			{
				std::wstring				uri;
				std::wstring				xml;	// the a:ext element
				std::vector<std::wstring>	rIds;	// its r:* attribute values, in document order
			};
			std::vector<UnknownExt>		UnknownExtLst;		// read from the document (pptx -> PPTY)
			std::vector<std::wstring>	UnknownExtLstXml;	// read from PPTY, relationships rewritten (PPTY -> pptx)

			void ReadUnknownExt(XmlUtils::CXmlNode& oNode);
			void WriteUnknownExtLstPPTY(NSBinPptxRW::CBinaryFileWriter* pWriter) const;
			void ReadUnknownExtLstPPTY(NSBinPptxRW::CBinaryFileReader* pReader);

			nullable_limit<Limit::BlipCompression> cstate;
			nullable<OOX::RId>	embed;
			nullable<OOX::RId>	link;
			std::wstring		m_namespace;

			std::wstring		imageFilepath;
	//internal
			std::wstring		mediaRid;
			std::wstring		mediaFilepath;
			bool				mediaExternal;

			std::wstring		oleRid;
			std::wstring		oleFilepathBin;
			std::wstring		oleFilepathImage;
			
			std::wstring		dataFilepathImage;
			std::string			dataFilepathImageA;

		protected:
			virtual void FillParentPointersForChilds();
		};
	} // namespace Logic
} // namespace PPTX
