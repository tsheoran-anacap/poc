import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class DatasetNormalization {
    public static void main(String[] args) {
        List<List<Double>> dataset = new ArrayList<>();
        dataset.add(Arrays.asList(10.0, 20.0, 30.0));
        dataset.add(Arrays.asList(55.0, 67.0, 90.0));
        dataset.add(Arrays.asList(100.0, 45.0, 98.0));

        List<List<Double>> normalizedDataset = normalizeDataset(dataset);

        // Print normalized dataset
        for (List<Double> point : normalizedDataset) {
            System.out.println(point);
        }
    }

    public static List<List<Double>> normalizeDataset(List<List<Double>> dataset) {
        int numDimensions = dataset.get(0).size();

        // Find the maximum and minimum values for each dimension
        double[] maxVals = new double[numDimensions];
        double[] minVals = new double[numDimensions];

        Arrays.fill(maxVals, Double.MIN_VALUE);
        Arrays.fill(minVals, Double.MAX_VALUE);

        for (List<Double> point : dataset) {
            for (int i = 0; i < numDimensions; i++) {
                if (point.get(i) > maxVals[i]) {
                    maxVals[i] = point.get(i);
                }

                if (point.get(i) < minVals[i]) {
                    minVals[i] = point.get(i);
                }
            }
        }

        // Normalize the dataset
        List<List<Double>> normalizedDataset = new ArrayList<>();

        for (List<Double> point : dataset) {
            List<Double> normalizedPoint = new ArrayList<>();

            for (int i = 0; i < numDimensions; i++) {
                double normalizedValue = (point.get(i) - minVals[i]) / (maxVals[i] - minVals[i]);
                normalizedPoint.add(Math.round(normalizedValue * 10.0) / 10.0); // Round to one decimal place
            }

            normalizedDataset.add(normalizedPoint);
        }

        return normalizedDataset;
    }
}
