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

#include "Blip.h"

#include "./../../Slide.h"
#include "./../../SlideMaster.h"
#include "./../../SlideLayout.h"
#include "./../../Theme.h"

#include "../../../../DesktopEditor/raster/ImageFileFormatChecker.h"
#include "../../../DocxFormat/Media/Image.h"
#include "../../../DocxFormat/FileTypes.h"

#include <set>

namespace PPTX
{
	namespace Logic
	{
		namespace
		{
			// a:blip/a:extLst/a:ext entries the model reads itself
			const wchar_t* c_sSvgBlipExtUri = L"{96DAC541-7B7A-43D3-8B79-37D633B846F1}";

			const wchar_t* c_sDrawingMlNs		= L"http://schemas.openxmlformats.org/drawingml/2006/main";
			const wchar_t* c_sRelationshipsNs	= L"http://schemas.openxmlformats.org/officeDocument/2006/relationships";

			// Every part this writes declares a: and r: on its root element. Other prefixes an
			// extension uses must be declared inside it, as PowerPoint does; these are added when a
			// document declared them further up.
			const wchar_t* c_arKnownExtNs[][2] =
			{
				{ L"a14",	L"http://schemas.microsoft.com/office/drawing/2010/main" },
				{ L"a15",	L"http://schemas.microsoft.com/office/drawing/2012/main" },
				{ L"a16",	L"http://schemas.microsoft.com/office/drawing/2014/main" },
				{ L"a1611",	L"http://schemas.microsoft.com/office/drawing/2016/11/main" },
				{ L"asvg",	L"http://schemas.microsoft.com/office/drawing/2016/SVG/main" },
				{ L"p14",	L"http://schemas.microsoft.com/office/powerpoint/2010/main" },
				{ L"p15",	L"http://schemas.microsoft.com/office/powerpoint/2012/main" }
			};

			// Parts an extension may relate: the HD Photo original of a corrected picture, or a picture.
			bool IsKeptExtPartType(const std::wstring& sRelsType)
			{
				return sRelsType == OOX::FileTypes::HdPhoto.RelationType() || sRelsType == OOX::FileTypes::Image.RelationType();
			}

			struct CExtXmlScan
			{
				std::set<std::wstring>		used;		// prefixes of element and attribute names
				std::set<std::wstring>		declared;	// prefixes declared inside the extension
				std::vector<std::wstring>	rIds;		// r:* attribute values
				bool						bad = false;
			};
			void ScanPrefix(const std::wstring& sName, CExtXmlScan& oScan)
			{
				size_t nPos = sName.find(L':');
				if (std::wstring::npos != nPos)
					oScan.used.insert(sName.substr(0, nPos));
			}
			// Writes an element as XML (empty elements self-closed), collecting its prefixes and
			// relationship ids; with pRIds, r:* values are replaced through it (an unknown one is bad).
			void WriteExtNode(XmlUtils::CXmlNode& oNode, NSStringUtils::CStringBuilder& oWriter, CExtXmlScan& oScan,
							  const std::map<std::wstring, std::wstring>* pRIds)
			{
				std::wstring sName = oNode.GetName();
				if (sName.empty())
				{
					oScan.bad = true;
					return;
				}
				ScanPrefix(sName, oScan);

				oWriter.AddCharSafe(wchar_t('<'));
				oWriter.WriteEncodeXmlString(sName);

				std::vector<std::wstring> arNames, arValues;
				oNode.ReadAllAttributes(arNames, arValues);
				for (size_t i = 0; i < arNames.size() && i < arValues.size(); ++i)
				{
					const std::wstring& sAttr = arNames[i];
					std::wstring sValue = arValues[i];

					if (0 == sAttr.find(L"xmlns:"))
					{
						std::wstring sPrefix = sAttr.substr(6);
						oScan.declared.insert(sPrefix);
						// r: must stay the relationships namespace, and no other prefix may name it
						if ((sPrefix == L"r") != (sValue == c_sRelationshipsNs))
							oScan.bad = true;
						if (sPrefix == L"a" && sValue != c_sDrawingMlNs)
							oScan.bad = true;
					}
					else if (sAttr != L"xmlns")
					{
						ScanPrefix(sAttr, oScan);
						if (0 == sAttr.find(L"r:"))
						{
							oScan.rIds.push_back(sValue);
							if (pRIds)
							{
								std::map<std::wstring, std::wstring>::const_iterator pFind = pRIds->find(sValue);
								if (pFind == pRIds->end())
									oScan.bad = true;
								else
									sValue = pFind->second;
							}
						}
					}
					oWriter.AddCharSafe(wchar_t(' '));
					oWriter.WriteEncodeXmlString(sAttr);
					oWriter.AddChar2Safe(wchar_t('='), wchar_t('\"'));
					oWriter.WriteEncodeXmlString(sValue);
					oWriter.AddCharSafe(wchar_t('\"'));
				}

				std::vector<XmlUtils::CXmlNode> arChildren;
				oNode.GetNodes(L"*", arChildren);
				std::wstring sText = oNode.GetText();

				if (arChildren.empty() && sText.empty())
				{
					oWriter.AddChar2Safe(wchar_t('/'), wchar_t('>'));
					return;
				}
				oWriter.AddCharSafe(wchar_t('>'));
				for (size_t i = 0; i < arChildren.size(); ++i)
					WriteExtNode(arChildren[i], oWriter, oScan, pRIds);
				oWriter.WriteEncodeXmlString(sText);
				oWriter.AddChar2Safe(wchar_t('<'), wchar_t('/'));
				oWriter.WriteEncodeXmlString(sName);
				oWriter.AddCharSafe(wchar_t('>'));
			}
			// An a:ext element as self-contained XML: its prefixes other than a: and r: declared
			// inside it. Empty if it cannot be (a prefix this does not know, a redeclared a: or r:).
			std::wstring GetExtXml(XmlUtils::CXmlNode& oNode, CExtXmlScan& oScan, const std::map<std::wstring, std::wstring>* pRIds)
			{
				if (oNode.GetName() != L"a:ext")
					return L"";

				NSStringUtils::CStringBuilder oWriter;
				WriteExtNode(oNode, oWriter, oScan, pRIds);
				if (oScan.bad)
					return L"";

				std::wstring sDeclare;
				for (std::set<std::wstring>::const_iterator it = oScan.used.begin(); it != oScan.used.end(); ++it)
				{
					const std::wstring& sPrefix = *it;
					if (sPrefix == L"a" || sPrefix == L"r" || sPrefix == L"xml" || oScan.declared.end() != oScan.declared.find(sPrefix))
						continue;

					bool bKnown = false;
					for (size_t i = 0; i < sizeof(c_arKnownExtNs) / sizeof(c_arKnownExtNs[0]); ++i)
					{
						if (sPrefix == c_arKnownExtNs[i][0])
						{
							sDeclare += L" xmlns:" + sPrefix + L"=\"" + c_arKnownExtNs[i][1] + L"\"";
							bKnown = true;
							break;
						}
					}
					if (!bKnown)
						return L"";
				}
				std::wstring sXml = oWriter.GetData();
				if (!sDeclare.empty())
					sXml.insert(6, sDeclare); // after "<a:ext"
				return sXml;
			}
		}

		Blip::Blip(std::wstring ns)
		{
			m_namespace = ns;
			mediaExternal = false;
		}
		Blip& Blip::operator=(const Blip& oSrc)
		{
			parentFile		= oSrc.parentFile;
			parentElement	= oSrc.parentElement;

			Effects = oSrc.Effects;

			UnknownExtLst		= oSrc.UnknownExtLst;
			UnknownExtLstXml	= oSrc.UnknownExtLstXml;

			cstate	= oSrc.cstate;
			embed	= oSrc.embed;
			link	= oSrc.link;

			m_namespace		= oSrc.m_namespace;

			oleRid			= oSrc.oleRid;
			oleFilepathBin	= oSrc.oleFilepathBin;

			mediaRid		= oSrc.mediaRid;
			mediaFilepath	= oSrc.mediaFilepath;
			mediaExternal	= oSrc.mediaExternal;

			return *this;
		}
		OOX::EElementType Blip::getType() const
		{
			return OOX::et_a_blip;
		}
		void Blip::ReadAttributes(XmlUtils::CXmlLiteReader& oReader)
		{
			WritingElement_ReadAttributes_Start_No_NS( oReader )
				WritingElement_ReadAttributes_Read_if     ( oReader, L"embed", embed)
				WritingElement_ReadAttributes_Read_else_if ( oReader, L"link", link )
				WritingElement_ReadAttributes_Read_else_if ( oReader, L"cstate", cstate )
			WritingElement_ReadAttributes_End_No_NS( oReader )
		}
		void Blip::fromXML(XmlUtils::CXmlLiteReader& oReader)
		{
			m_namespace = XmlUtils::GetNamespace(oReader.GetName());
			
			ReadAttributes( oReader );

			Effects.clear();
			if ( oReader.IsEmptyNode() )
				return;
				
			int nParentDepth = oReader.GetDepth();
			while( oReader.ReadNextSiblingNode( nParentDepth ) )
			{
				std::wstring strName = XmlUtils::GetNameNoNS(oReader.GetName());

				if (strName == L"extLst")
				{
					if (oReader.IsEmptyNode())
						continue;

					int nParentDepth1 = oReader.GetDepth();
					while (oReader.ReadNextSiblingNode(nParentDepth1))
					{
						Ext element;
						ExtLst.push_back(element);
						ExtLst.back().fromXML(oReader);
					}
				}
				else
				{
					UniEffect uni;
					Effects.push_back(uni);
					Effects.back().fromXML(oReader);
				}
			}
		
			FillParentPointersForChilds();
		}
		void Blip::fromXML(XmlUtils::CXmlNode& node)
		{
			m_namespace = XmlUtils::GetNamespace(node.GetName());

            XmlMacroReadAttributeBase(node, L"r:embed", embed);
            XmlMacroReadAttributeBase(node, L"r:link", link);
            XmlMacroReadAttributeBase(node, L"cstate", cstate);

			XmlUtils::CXmlNode list = node.ReadNodeNoNS(L"extLst");
			if (list.IsValid())
			{
				std::vector<XmlUtils::CXmlNode> oNodes;
				if (list.GetNodes(L"*", oNodes))
				{
					size_t nCount = oNodes.size();
					for (size_t i = 0; i < nCount; ++i)
					{
						XmlUtils::CXmlNode& oNode = oNodes[i];

						Ext element;
						element.fromXML(oNode);

						ExtLst.push_back(element);

						ReadUnknownExt(oNode);
					}
				}
			}
			Effects.clear();
			XmlMacroLoadArray(node, L"*", Effects, UniEffect);
			FillParentPointersForChilds();
		}
		std::wstring Blip::toXML() const
		{
			XmlUtils::CAttribute oAttr;

			if (embed.IsInit())
				oAttr.Write(L"r:embed", embed->ToString());
			if (link.IsInit())
				oAttr.Write(L"r:link", link->ToString());
			oAttr.WriteLimitNullable(L"cstate", cstate);

			XmlUtils::CNodeValue oValue;
			oValue.WriteArray(Effects);

			std::wstring strName = m_namespace.empty() ? L"blip" : m_namespace + L":blip";
			return XmlUtils::CreateNode(strName, oAttr, oValue);
		}
		void Blip::ReadUnknownExt(XmlUtils::CXmlNode& oNode)
		{
			std::wstring sUri = oNode.GetAttribute(L"uri");
			if (sUri.empty() || sUri == c_sSvgBlipExtUri)
				return;

			CExtXmlScan oScan;
			std::wstring sXml = GetExtXml(oNode, oScan, NULL);
			if (sXml.empty())
				return;

			UnknownExt oExt;
			oExt.uri	= sUri;
			oExt.xml	= sXml;
			oExt.rIds	= oScan.rIds;
			UnknownExtLst.push_back(oExt);
		}
		void Blip::WriteUnknownExtLstPPTY(NSBinPptxRW::CBinaryFileWriter* pWriter) const
		{
			if (UnknownExtLst.empty())
				return;

			// presentation parts only (slides, layouts, masters, notes, themes): their rels are this
			// document's, and the binary -> pptx side writes the parts back into ppt/media
			OOX::IFileContainer* pContainer = NULL;
			OOX::IFileContainer* pRels = pWriter->GetRelsPtr();
			if (pRels)
				pContainer = dynamic_cast<PPTX::FileContainer*>(pRels);
			else if (parentFileIs<FileContainer>())
				pContainer = dynamic_cast<OOX::IFileContainer*>(&parentFileAs<FileContainer>());
			if (!pContainer || !pWriter->m_pCommon || !pWriter->m_pCommon->m_pMediaManager)
				return;

			struct CPart { std::wstring rId, relsType, path; };
			struct CEntry { const UnknownExt* ext; std::vector<CPart> parts; };
			std::vector<CEntry> arEntries;

			for (size_t i = 0; i < UnknownExtLst.size(); ++i)
			{
				const UnknownExt& oExt = UnknownExtLst[i];
				CEntry oEntry;
				oEntry.ext = &oExt;

				bool bOk = true;
				for (size_t j = 0; j < oExt.rIds.size() && bOk; ++j)
				{
					bool bSeen = false;
					for (size_t k = 0; k < oEntry.parts.size(); ++k)
						if (oEntry.parts[k].rId == oExt.rIds[j]) bSeen = true;
					if (bSeen) continue;

					smart_ptr<OOX::Media> pMedia = pContainer->Find(OOX::RId(oExt.rIds[j])).smart_dynamic_cast<OOX::Media>();
					if (!pMedia.IsInit() || pMedia->IsExternal() || !pMedia->IsExist())
					{
						bOk = false;
						break;
					}
					CPart oPart;
					oPart.rId		= oExt.rIds[j];
					oPart.relsType	= pMedia->type().RelationType();
					if (!IsKeptExtPartType(oPart.relsType))
					{
						bOk = false;
						break;
					}
					NSShapeImageGen::CMediaInfo oInfo = pWriter->m_pCommon->m_pMediaManager->WriteMedia(pMedia->filename().GetPath());
					oPart.path = oInfo.GetPath2();
					oEntry.parts.push_back(oPart);
				}
				if (bOk)
					arEntries.push_back(oEntry);
			}
			if (arEntries.empty())
				return;

			// record 4: ULONG count, then per a:ext a record 0 with string attributes 0 uri, 1 xml and
			// a record 0 per related part with string attributes 0 rId (as in the xml), 1 relationship
			// type, 2 file name in the media folder. Readers that do not know record 4 skip it.
			pWriter->StartRecord(4);
			pWriter->WriteULONG((ULONG)arEntries.size());
			for (size_t i = 0; i < arEntries.size(); ++i)
			{
				pWriter->StartRecord(0);
				pWriter->WriteBYTE(NSBinPptxRW::g_nodeAttributeStart);
				pWriter->WriteString1(0, arEntries[i].ext->uri);
				pWriter->WriteString1(1, arEntries[i].ext->xml);
				pWriter->WriteBYTE(NSBinPptxRW::g_nodeAttributeEnd);

				for (size_t j = 0; j < arEntries[i].parts.size(); ++j)
				{
					const CPart& oPart = arEntries[i].parts[j];
					pWriter->StartRecord(0);
					pWriter->WriteBYTE(NSBinPptxRW::g_nodeAttributeStart);
					pWriter->WriteString1(0, oPart.rId);
					pWriter->WriteString1(1, oPart.relsType);
					pWriter->WriteString1(2, oPart.path);
					pWriter->WriteBYTE(NSBinPptxRW::g_nodeAttributeEnd);
					pWriter->EndRecord();
				}
				pWriter->EndRecord();
			}
			pWriter->EndRecord();
		}
		void Blip::ReadUnknownExtLstPPTY(NSBinPptxRW::CBinaryFileReader* pReader)
		{
			// at record 4's length (see WriteUnknownExtLstPPTY)
			LONG _end = pReader->GetPos() + pReader->GetLong() + 4;

			bool bPresentation = pReader->m_pRels && pReader->m_pRels->m_pManager &&
				pReader->m_pRels->m_pManager->m_nDocumentType == XMLWRITER_DOC_TYPE_PPTX;

			ULONG nCount = pReader->GetULong();
			for (ULONG i = 0; i < nCount && pReader->GetPos() < _end; ++i)
			{
				pReader->Skip(1); // type 0
				LONG _e = pReader->GetPos() + pReader->GetLong() + 4;

				std::wstring sXml;
				pReader->Skip(1); // start attributes
				while (true)
				{
					BYTE _at = pReader->GetUChar_TypeNode();
					if (NSBinPptxRW::g_nodeAttributeEnd == _at)
						break;
					std::wstring sValue = pReader->GetString2();
					if (1 == _at)
						sXml = sValue;
				}

				struct CPart { std::wstring rId, relsType, path; };
				std::vector<CPart> arParts;
				while (pReader->GetPos() < _e)
				{
					BYTE _t = pReader->GetUChar();
					if (0 != _t)
					{
						pReader->SkipRecord();
						continue;
					}
					LONG _e1 = pReader->GetPos() + pReader->GetLong() + 4;
					CPart oPart;
					pReader->Skip(1); // start attributes
					while (true)
					{
						BYTE _at = pReader->GetUChar_TypeNode();
						if (NSBinPptxRW::g_nodeAttributeEnd == _at)
							break;
						std::wstring sValue = pReader->GetString2();
						if (0 == _at)		oPart.rId		= sValue;
						else if (1 == _at)	oPart.relsType	= sValue;
						else if (2 == _at)	oPart.path		= sValue;
					}
					pReader->Seek(_e1);
					arParts.push_back(oPart);
				}
				pReader->Seek(_e);

				// only with the picture itself written (record 3 comes first)
				if (!bPresentation || sXml.empty() || !embed.IsInit())
					continue;

				XmlUtils::CXmlNode oRoot;
				if (!oRoot.FromXmlString(std::wstring(L"<w xmlns:a=\"") + c_sDrawingMlNs + L"\" xmlns:r=\"" + c_sRelationshipsNs + L"\">" + sXml + L"</w>"))
					continue;
				std::vector<XmlUtils::CXmlNode> arExt;
				oRoot.GetNodes(L"*", arExt);
				if (1 != arExt.size())
					continue;

				// every relationship the extension uses must come with its part, and every part must
				// be a file in the media folder; otherwise drop the extension, relating nothing
				std::map<std::wstring, std::wstring> mapFiles; // rId -> file
				std::map<std::wstring, std::wstring> mapTypes; // rId -> relationship type
				bool bOk = true;
				std::wstring sMediaFolder = OOX::CPath(pReader->m_strFolder + FILE_SEPARATOR_STR + L"media").GetPath();
				for (size_t j = 0; j < arParts.size() && bOk; ++j)
				{
					const CPart& oPart = arParts[j];
					const std::wstring& sPath = oPart.path;
					if (oPart.rId.empty() || !IsKeptExtPartType(oPart.relsType) || sPath.empty() ||
						0 == sPath.find(L"http:") || 0 == sPath.find(L"https:") || 0 == sPath.find(L"ftp:") ||
						0 == sPath.find(L"file:") || 0 == sPath.find(L"data:") || 0 == sPath.find(L"theme"))
					{
						bOk = false;
						break;
					}
					std::wstring sFile = OOX::CPath(sMediaFolder + FILE_SEPARATOR_STR + sPath).GetPath();
					if (0 != sFile.find(sMediaFolder + FILE_SEPARATOR_STR) || !NSFile::CFileBinary::Exists(sFile))
					{
						bOk = false;
						break;
					}
					mapFiles[oPart.rId] = sFile;
					mapTypes[oPart.rId] = oPart.relsType;
				}
				if (!bOk)
					continue;

				CExtXmlScan oCheck;
				if (GetExtXml(arExt[0], oCheck, NULL).empty())
					continue;
				for (size_t j = 0; j < oCheck.rIds.size(); ++j)
					if (mapFiles.end() == mapFiles.find(oCheck.rIds[j]))
						bOk = false;
				if (!bOk)
					continue;

				std::map<std::wstring, std::wstring> mapRIds; // old rId -> new rId, for the parts the xml uses
				for (size_t j = 0; j < oCheck.rIds.size() && bOk; ++j)
				{
					const std::wstring& sRId = oCheck.rIds[j];
					if (mapRIds.end() != mapRIds.find(sRId))
						continue;
					NSBinPptxRW::_relsGeneratorInfo oInfo = pReader->m_pRels->WriteMediaPart(mapFiles[sRId], mapTypes[sRId]);
					if (0 == oInfo.nMediaRId)
						bOk = false;
					else
						mapRIds[sRId] = L"rId" + std::to_wstring(oInfo.nMediaRId);
				}
				if (!bOk)
					continue;

				CExtXmlScan oScan;
				std::wstring sOut = GetExtXml(arExt[0], oScan, &mapRIds);
				if (!sOut.empty())
					UnknownExtLstXml.push_back(sOut);
			}
			pReader->Seek(_end);
		}
		void Blip::FillParentPointersForChilds()
		{
			for (size_t i = 0; i < Effects.size(); ++i)
			{
				Effects[i].SetParentPointer(this);
			}

			for (size_t i = 0; i < ExtLst.size(); ++i)
			{
				ExtLst[i].SetParentPointer(this);
			}

		}
		std::wstring Blip::GetFullPicName(OOX::IFileContainer* pRels)const
		{
			std::wstring result;
			for (size_t i = 0; i < ExtLst.size(); ++i)
			{
				if (ExtLst[i].link_svg.IsInit())
				{
					if (pRels != NULL)
					{
						smart_ptr<OOX::Image> p = pRels->Get<OOX::Image>(*ExtLst[i].link_svg);
						if (p.is_init())
						{
							result = p->filename().m_strFilename;
							break;
						}
					}

					if (parentFileIs<FileContainer>())
					{
						result = parentFileAs<FileContainer>().GetImagePathNameFromRId(*ExtLst[i].link_svg);
						break;
					}
				}
			}
			if (false == result.empty()) return result;
			
			if (embed.IsInit())
			{
				if (pRels != NULL)
				{
					smart_ptr<OOX::Image> p = pRels->Get<OOX::Image>(*embed);
					if (p.is_init())
						return p->filename().m_strFilename;
				}

				if (parentFileIs<FileContainer>())
					return parentFileAs<FileContainer>().GetImagePathNameFromRId(*embed);

				return L"";
			}
			else if (link.IsInit())
			{
				if (pRels != NULL)
				{
					smart_ptr<OOX::Image> p = pRels->Get<OOX::Image>(*link);
					if (p.is_init())
						return p->filename().m_strFilename;
				}

				if (parentFileIs<FileContainer>())
					return parentFileAs<FileContainer>().GetImagePathNameFromRId(*link);

				return L"";
			}
			return L"";
		}
		std::wstring Blip::GetFullOleName(const OOX::RId& oRId, OOX::IFileContainer* pRels)const
		{
			smart_ptr<OOX::OleObject> pOleObject;
			
			if (pRels != NULL)
				pOleObject = pRels->Get<OOX::OleObject>(oRId);
			else
			{
				OOX::IFileContainer* pContainer = dynamic_cast<OOX::IFileContainer*>(const_cast<PPTX::WrapperFile*>(parentFile));
				if (pContainer)
					pOleObject = pContainer->Get<OOX::OleObject>(oRId);
			}
			
			if (pOleObject.IsInit())
				return pOleObject->filename().m_strFilename;
			
			return L"";
		}
		void Blip::toXmlWriter(NSBinPptxRW::CXmlWriter* pWriter) const
		{
			std::wstring strName;

			if (XMLWRITER_DOC_TYPE_WORDART == pWriter->m_lDocType)
				strName = L"w14:blip";
			else
				strName = m_namespace.empty() ? L"blip" : (m_namespace + L":blip");

			pWriter->StartNode(strName);

			pWriter->StartAttributes();
			if (embed.IsInit())
				pWriter->WriteAttribute(L"r:embed", embed->ToString());
			if (link.IsInit())
				pWriter->WriteAttribute(L"r:link", link->ToString());
			pWriter->WriteAttribute(L"cstate", cstate);
			pWriter->EndAttributes();

			for (size_t i = 0; i < Effects.size(); ++i)
			{
				Effects[i].toXmlWriter(pWriter);
			}
			if (UnknownExtLstXml.empty())
			{
				pWriter->WriteArray(L"a:extLst", ExtLst);
			}
			else
			{
				pWriter->StartNode(L"a:extLst");
				pWriter->EndAttributes();
				for (size_t i = 0; i < UnknownExtLstXml.size(); ++i)
					pWriter->WriteString(UnknownExtLstXml[i]);
				for (size_t i = 0; i < ExtLst.size(); ++i)
					ExtLst[i].toXmlWriter(pWriter);
				pWriter->EndNode(L"a:extLst");
			}

			pWriter->EndNode(strName);
		}
		void Blip::toPPTY(NSBinPptxRW::CBinaryFileWriter* pWriter) const
		{
			pWriter->WriteBYTE(NSBinPptxRW::g_nodeAttributeStart);
			pWriter->WriteLimit2(0, cstate);
			pWriter->WriteBYTE(NSBinPptxRW::g_nodeAttributeEnd);

			if (embed.is_init())
				pWriter->WriteString1(10, embed->get());
			if (link.is_init())
				pWriter->WriteString1(11, link->get());

			pWriter->StartRecord(2);
			ULONG len = (ULONG)Effects.size();
			pWriter->WriteULONG(len);

			for (ULONG i = 0; i < len; ++i)
			{
				pWriter->WriteRecord1(0, Effects[i]);// id isn't important
			}

			pWriter->EndRecord();

			double dW = pWriter->GetShapeWidth() / 12700. ; //emu to pt 
			double dH = pWriter->GetShapeHeight() / 12700.;

			OOX::IFileContainer* pRels = pWriter->GetRelsPtr();

			std::wstring	additionalPath;
			int				additionalType = 0;
			if (!oleFilepathBin.empty())
			{
				additionalPath = oleFilepathBin;
				additionalType = 1;
			}
			else if (!oleRid.empty())
			{
				additionalPath	= this->GetFullOleName(OOX::RId(oleRid), pRels);
				additionalType	= 1;
			}
			else if (!mediaRid.empty())
			{
				additionalPath	= this->GetFullOleName(OOX::RId(mediaRid), pRels);
				additionalType	= 2;
			}
			else if (!mediaFilepath.empty())
			{
				additionalPath	= mediaFilepath;
				additionalType	= 2;
			}
			NSShapeImageGen::CMediaInfo oId;
			if (!dataFilepathImageA.empty())
			{
				oId = pWriter->m_pCommon->m_pMediaManager->WriteImage(dataFilepathImageA, dW, dH, additionalPath, additionalType);
			}
			else if (!dataFilepathImage.empty())
			{
				oId = pWriter->m_pCommon->m_pMediaManager->WriteImage(dataFilepathImage, dW, dH, additionalPath, additionalType);
			}
			else if (!oleFilepathImage.empty())
			{
				std::wstring imagePath = oleFilepathImage;
				oId = pWriter->m_pCommon->m_pMediaManager->WriteImage(imagePath, dW, dH, additionalPath, additionalType);
			}
			else
			{
				std::wstring imagePath = this->GetFullPicName(pRels);
				oId = pWriter->m_pCommon->m_pMediaManager->WriteImage(imagePath, dW, dH, additionalPath, additionalType);
			}

			std::wstring s = oId.GetPath2();

			pWriter->StartRecord(3);

			pWriter->WriteBYTE(NSBinPptxRW::g_nodeAttributeStart);
			pWriter->WriteString1(0, s);
			pWriter->WriteBYTE(NSBinPptxRW::g_nodeAttributeEnd);

			pWriter->EndRecord();

			WriteUnknownExtLstPPTY(pWriter);
		}
		void Blip::fromPPTY(NSBinPptxRW::CBinaryFileReader* pReader)
		{
			LONG _s2 = pReader->GetPos();
			LONG _e2 = _s2 + pReader->GetLong() + 4;

			pReader->Skip(1);

			while (true)
			{
				BYTE _at = pReader->GetUChar_TypeNode();
				if (NSBinPptxRW::g_nodeAttributeEnd == _at)
					break;

				if (_at == 0)
					pReader->Skip(1);
			}

			while (pReader->GetPos() < _e2)
			{
				BYTE _t = pReader->GetUChar();

				switch (_t)
				{
					case 0:
					case 1:
					{
						// id. embed / link
						pReader->Skip(4);						
					}break;
					case 10:
					case 11:
					{
						// id. embed / link
						pReader->GetString2();						
					}break;
					case 2:
					{
						pReader->Skip(4);
						ULONG count_effects = pReader->GetULong();
						for (ULONG _eff = 0; _eff < count_effects; ++_eff)
						{
							pReader->Skip(1); // type 

							Effects.push_back(UniEffect());
							Effects.back().fromPPTY(pReader);

							if (false == Effects.back().is_init())
							{
								Effects.pop_back();
							}
						}						
					}break;
					case 3:
					{
						pReader->Skip(6); // len + start attributes + type

						std::wstring strImagePath = pReader->GetString2(true);

						if (0 != strImagePath.find(L"http:") &&
							0 != strImagePath.find(L"https:") &&
							0 != strImagePath.find(L"ftp:") &&
							0 != strImagePath.find(L"file:"))
						{
							OOX::CPath pathNormalizer;

							if (0 == strImagePath.find(L"theme"))
							{
								pathNormalizer = pReader->m_strFolderExternalThemes;
							}
							else
							{
								pathNormalizer = pReader->m_strFolder + FILE_SEPARATOR_STR + L"media";
							}
							std::wstring strPath = pathNormalizer.GetPath();

							strImagePath = strPath  + FILE_SEPARATOR_STR + strImagePath;

							pathNormalizer = strImagePath;
							strImagePath = pathNormalizer.GetPath();

							if (std::wstring::npos != strImagePath.find(strPath))
							{
								CImageFileFormatChecker checker;
								if (false == checker.isImageFile(strImagePath))
								{
									strImagePath.clear();
								}				
							}
							else strImagePath.clear();
						}	
						if (false == strImagePath.empty())
						{
							std::vector<smart_ptr<OOX::File>> additionalFiles;
							NSBinPptxRW::_relsGeneratorInfo oRelsGeneratorInfo = pReader->m_pRels->WriteImage(strImagePath, additionalFiles, L"", L"");

							if (oRelsGeneratorInfo.nImageRId > 0)
							{
								embed = new OOX::RId(oRelsGeneratorInfo.nImageRId);
							}
						}
						pReader->Skip(1); // end attribute
					}break;
					case 4:
					{
						ReadUnknownExtLstPPTY(pReader);
					}break;
					default:
					{
						pReader->SkipRecord();
						break;
					}
				}
			}

			pReader->Seek(_e2);
		}
	} // namespace Logic
} // namespace PPTX
