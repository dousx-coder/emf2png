package cn.cruder;

import net.arnx.wmf2svg.gdi.svg.SvgGdi;
import net.arnx.wmf2svg.gdi.wmf.WmfParser;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.freehep.graphicsio.emf.EMFInputStream;
import org.freehep.graphicsio.emf.EMFRenderer;
import org.w3c.dom.Document;

import javax.imageio.ImageIO;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.List;
import java.util.*;
import java.util.zip.GZIPOutputStream;

/**
 * @Author: cruder
 * @Date: 2021/09/13/16:08
 */
public class EmfUtils {
    /**
     * 获得文档的表格
     * path 文件路径
     *
     * @return
     */
    public static List<Map<String, Object>> getTables(String path) {
        try {
            FileInputStream input = new FileInputStream(path);
            POIFSFileSystem pfs = new POIFSFileSystem(input);
            HWPFDocument hwpf = new HWPFDocument(pfs);
            // 返回范围涵盖整个文档,但不包括任何页眉和页脚。
            Range range = hwpf.getRange();
            // 表迭代器
            TableIterator it = new TableIterator(range);
            int tableNum = 1;
            List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
            while (it.hasNext()) {
                StringBuffer sb = new StringBuffer();
                // 遍历当前表格
                Table tb = (Table) it.next();
                String tableStart = "<table border=\"1\" cellspacing=\"0\"><center>";
                sb.append(tableStart);
                for (int i = 0; i < tb.numRows(); i++) {
                    // 获取当前表格的行
                    TableRow tr = tb.getRow(i);
                    String trStart = "<tr>";
                    sb.append(trStart);
                    // 获取行的内部细胞
                    for (int j = 0; j < tr.numCells(); j++) {
                        // 获取表格单元格
                        TableCell td = tr.getCell(j);
                        // 用于获取段落的数量在一个范围内。如果这个范围小于一个段落,它将返回1包含段落。
                        for (int k = 0; k < td.numParagraphs(); k++) {
                            // 段落在这个范围内的指定索引。
                            Paragraph p = td.getParagraph(k);
                            // 遍历当前td的内容
                            String text = p.text().trim();
                            String tds = null;
                            if (i == 0) {
                                tds = "<td><b>" + text + "</b></td>";
                            } else {
                                tds = "<td>" + text + "</td>";
                            }
                            sb.append(tds);
                        }
                    }
                    String trEnd = "</tr>";
                    sb.append(trEnd);
                }
                String tableEnd = "</center></table>";
                sb.append(tableEnd);
                Map<String, Object> m = new HashMap<String, Object>();
                m.put("" + tableNum + "", sb.toString());
                tableNum++;
                if (!list.contains(m)) {
                    list.add(m);
                }
            }
            if (hwpf != null) {
                hwpf.close();
            }
            if (input != null) {
                input.close();
            }
            return list;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 获取图片的路径
     *
     * @param path 文档的路径
     * @return 返回list集合，里面存放的是图片的集合
     */
    public static List<Map<String, Object>> getWordImageUrl(String path) {
        File file = null;
        List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
        List<String> emfs = new ArrayList<String>();
        try {
            file = new File(path);
            FileInputStream f = new FileInputStream(file.getAbsolutePath());
            HWPFDocument doc = new HWPFDocument(f);
            // 获取图片表
            PicturesTable pTable = doc.getPicturesTable();
            // 返回字符长度
            int length = doc.characterLength();
            // 文档目录，用于删除emf文件的时候传入路径
            String directory = "";
            if (length > 0) {
                //String[] ym = getYearAndMonth();// 获得年+月
                // 如果二级目录没有的话则生成
                directory = "C:\\Users\\dousx\\temp\\";
                File fp1 = new File(directory);
                if (!fp1.exists() && !fp1.isDirectory()) {
                    fp1.mkdirs();
                }
                for (int i = 0; i < length; i++) {
                    Range range = new Range(i, i + 1, doc);
                    // 得到这个角色在索引。
                    CharacterRun cr = range.getCharacterRun(0);
                    // 确定指定字符运行包含参考图片
                    if (pTable.hasPicture(cr)) {
                        // 将遍历到的图片进行解析,生成emf图片,并保存到磁盘中
                        Picture pic = pTable.extractPicture(cr, false);
                        String afileName = pic.suggestFullFileName();
                        afileName = afileName.substring(afileName.length() - 4);
                        String fileMainName = UUID.randomUUID().toString()
                                .replace("-", "");
                        //String saveUrl = "G:/upload/img/" + ym[0] + "/" + ym[1]
                        //+ "/" + fileMainName + afileName;
                        String saveUrl = "C:\\Users\\dousx\\temp\\" + fileMainName + afileName;
                        // 将文件读取到输出流中
                        OutputStream out = new FileOutputStream(new File(
                                saveUrl));
                        // 写入到磁盘中
                        pic.writeImageContent(out);
                        if (out != null) {
                            out.close();
                        }
                        if (!emfs.contains(saveUrl)) {
                            emfs.add(saveUrl);
                        }
                    }
                }
            }
            if (doc != null) {
                doc.close();
            }
            if (f != null) {
                f.close();
            }
            // 将emf转为png
            emfConversionPng(emfs);
            return list;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    /**
     * emf或者wmf转换为png图片格式
     *
     * @param list
     * @return
     * @throws IOException
     */
    public static void emfConversionPng(List<String> list) throws IOException {
        if (list.size() > 0) {
            // 对文件的命名进行重新修改
            for (int i = 0; i < list.size(); i++) {
                String saveUrl = list.get(i);
                // 从doc文档解析的图片很有可能已经是png了，所以此处需要判断
                if (saveUrl.contains("emf") || saveUrl.contains("EMF")) {
                    InputStream is = new FileInputStream(saveUrl);
                    EMFInputStream eis = new EMFInputStream(is, EMFInputStream.DEFAULT_VERSION);
                    EMFRenderer emfRenderer = new EMFRenderer(eis);
                    final int width = (int) eis.readHeader().getBounds().getWidth();
                    final int height = (int) eis.readHeader().getBounds().getHeight();
                    // 设置图片的大小和样式
                    final BufferedImage result = new BufferedImage(width + 60, height + 40, BufferedImage.TYPE_4BYTE_ABGR);
                    Graphics2D g2 = (Graphics2D) result.createGraphics();
                    emfRenderer.paint(g2);
                    String url = saveUrl.replace(saveUrl.substring(saveUrl.length() - 3), "png");
                    File outputfile = new File(url);
                    // 写入到磁盘中(格式设置为png背景不会变为橙色)
                    ImageIO.write(result, "png", outputfile);
                    // 当前的图片写入到磁盘中后，将流关闭
                    if (eis != null) {
                        eis.close();
                    }
                    if (is != null) {
                        is.close();
                    }
                } else if (saveUrl.contains("wmf") || saveUrl.contains("WMF")) {
                    // 将wmf转svg
                    String svgFile = saveUrl.substring(0, saveUrl.lastIndexOf(".wmf"))
                            + ".svg";
                    wmfToSvg(saveUrl, svgFile);
                    // 将svg转png
                    String jpgFile = saveUrl.substring(0, saveUrl.lastIndexOf(".wmf"))
                            + ".png";
                    // TODO: 2021-09-13
                    //svgToJpg(svgFile, jpgFile);
                }
            }
        }
    }

    /**
     * 将wmf转换为svg
     *
     * @param src
     * @param dest
     */
    public static void wmfToSvg(String src, String dest) {
        File file = new File(src);
        boolean compatible = false;
        try {
            InputStream in = new FileInputStream(file);
            WmfParser parser = new WmfParser();
            final SvgGdi gdi = new SvgGdi(compatible);
            parser.parse(in, gdi);

            Document doc = gdi.getDocument();
            OutputStream out = new FileOutputStream(dest);
            if (dest.endsWith(".svgz")) {
                out = new GZIPOutputStream(out);
            }
            output(doc, out);
            if (out != null) {
                out.close();
            }
            if (in != null) {
                in.close();
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {

        }
    }

    /**
     * 输出信息
     *
     * @param doc
     * @param out
     * @throws Exception
     */
    private static void output(Document doc, OutputStream out) throws Exception {
        TransformerFactory factory = TransformerFactory.newInstance();
        Transformer transformer = factory.newTransformer();
        transformer.setOutputProperty(OutputKeys.METHOD, "xml");
        transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        transformer.setOutputProperty(OutputKeys.DOCTYPE_PUBLIC,
                "-//W3C//DTD SVG 1.0//EN");
        transformer.setOutputProperty(OutputKeys.DOCTYPE_SYSTEM,
                "http://www.w3.org/TR/2001/REC-SVG-20010904/DTD/svg10.dtd");
        transformer.transform(new DOMSource(doc), new StreamResult(out));
        if (out != null) {
            out.flush();
            out.close();
        }
    }


    /**
     * 将svg转化为JPG
     *
     * @param src
     * @param dest
     */
    //public static void svgToJpg(String src, String dest) {
    //    FileOutputStream jpgOut = null;
    //    FileInputStream svgStream = null;
    //    ByteArrayOutputStream svgOut = null;
    //    ByteArrayInputStream svgInputStream = null;
    //    ByteArrayOutputStream jpg = null;
    //    File svg = null;
    //    try {
    //        // 获取到svg文件
    //        svg = new File(src);
    //        svgStream = new FileInputStream(svg);
    //        svgOut = new ByteArrayOutputStream();
    //        // 获取到svg的stream
    //        int noOfByteRead = 0;
    //        while ((noOfByteRead = svgStream.read()) != -1) {
    //            svgOut.write(noOfByteRead);
    //        }
    //        ImageTranscoder it = new PNGTranscoder();
    //        it.addTranscodingHint(JPEGTranscoder.KEY_QUALITY, new Float(1f));
    //        it.addTranscodingHint(ImageTranscoder.KEY_WIDTH, new Float(500));
    //        jpg = new ByteArrayOutputStream();
    //        svgInputStream = new ByteArrayInputStream(svgOut.toByteArray());
    //        it.transcode(new TranscoderInput(svgInputStream),
    //                new TranscoderOutput(jpg));
    //        jpgOut = new FileOutputStream(dest);
    //        jpgOut.write(jpg.toByteArray());
    //    } catch (Exception e) {
    //        e.printStackTrace();
    //    } finally {
    //        try {
    //            if (svgInputStream != null) {
    //                svgInputStream.close();
    //            }
    //            if (jpg != null) {
    //                jpg.close();
    //            }
    //            if (svgStream != null) {
    //                svgStream.close();
    //
    //            }
    //            if (svgOut != null) {
    //                svgOut.close();
    //            }
    //            if (jpgOut != null) {
    //                jpgOut.flush();
    //                jpgOut.close();
    //            }
    //            if (svg != null) {
    //                svg.delete();
    //            }
    //        } catch (IOException e) {
    //            e.printStackTrace();
    //        }
    //    }
    //}


    /**
     * 获取文件路径
     *
     * @param dirName  文件夹
     * @param fileName 文件名
     * @return
     * @throws IOException
     */
    private static String getFilePath(String dirName, String fileName) throws IOException {
        String resourcesPath = "src" + File.separator + "main" + File.separator + "resources";
        File directory = new File(resourcesPath);
        String courseFile = directory.getCanonicalPath();
        return courseFile + File.separator + dirName + File.separator + fileName;

    }



    public static void main(String[] args) throws IOException {

        String excelPath = getFilePath("emf", "81b7d45d-e25e-43fc-b7cb-79b437306f6d.emf");
        LinkedList<String> strings = new LinkedList<>();
        strings.add(excelPath);
        emfConversionPng(strings);
    }



}
